"""
Little Helper - Monitor server.

Serves hardware monitor snapshots over HTTP and WebSocket using Starlette.
Optionally broadcasts the service via mDNS/DNS-SD for easy client discovery.
"""

import asyncio
from contextlib import asynccontextmanager
import logging
import socket
import threading

import system_overlay

log = logging.getLogger("little_helper.monitor_server")

_STARLETTE_IMPORT_ERROR = None

try:
    from starlette.applications import Starlette
    from starlette.responses import JSONResponse
    from starlette.routing import Route, WebSocketRoute
    from starlette.websockets import WebSocketDisconnect
    import uvicorn
    import websockets  # noqa: F401
except Exception as exc:
    Starlette = None
    JSONResponse = None
    Route = None
    WebSocketRoute = None
    WebSocketDisconnect = Exception
    uvicorn = None
    _STARLETTE_IMPORT_ERROR = exc

_ZEROCONF_IMPORT_ERROR = None

try:
    from zeroconf import Zeroconf, ServiceInfo
except Exception as exc:
    Zeroconf = None
    ServiceInfo = None
    _ZEROCONF_IMPORT_ERROR = exc


DEFAULT_HOST = "0.0.0.0"
DEFAULT_PORT = 9980
DEFAULT_WS_INTERVAL_MS = 1000
MIN_WS_INTERVAL_MS = 200
MAX_WS_INTERVAL_MS = 60000
MDNS_SERVICE_TYPE = "_lhm._tcp.local."
MDNS_WATCH_INTERVAL_SEC = 10.0

_MDNS_172_16_NET = None  # lazily initialised IPv4Network for 172.16/12


def monitor_server_dependencies_available() -> tuple[bool, str | None]:
    if _STARLETTE_IMPORT_ERROR is None:
        return True, None
    return False, str(_STARLETTE_IMPORT_ERROR)


def normalize_monitor_server_config(config: dict) -> dict:
    raw_cfg = config.get("monitor_server", {})
    host = str(raw_cfg.get("host", DEFAULT_HOST)).strip() or DEFAULT_HOST
    token = str(raw_cfg.get("token", "")).strip()
    try:
        port = int(raw_cfg.get("port", DEFAULT_PORT))
    except (TypeError, ValueError):
        port = DEFAULT_PORT
    port = max(1, min(65535, port))
    return {
        "enabled": bool(raw_cfg.get("enabled", False)),
        "host": host,
        "port": port,
        "token": token,
        "mdns": bool(raw_cfg.get("mdns", True)),
    }


def get_monitor_urls(server_cfg: dict) -> dict:
    host = server_cfg.get("host", DEFAULT_HOST)
    port = server_cfg.get("port", DEFAULT_PORT)
    display_host = "127.0.0.1" if host == "0.0.0.0" else host
    return {
        "http": f"http://{display_host}:{port}/api/monitor",
        "websocket": f"ws://{display_host}:{port}/ws/monitor",
    }


def _get_local_ip() -> str | None:
    """Pick the best local IPv4 for mDNS, or None if none is usable.

    Priority (lower wins):
      0: 192.168.x.x         (typical home/office LAN)
      1: 10.x.x.x            (corporate LAN)
      2: other private ranges
      3: public addresses
    Skips loopback, APIPA (169.254/16), and 172.16/12 (usually Docker/WSL/Hyper-V).
    """
    import ipaddress

    global _MDNS_172_16_NET
    if _MDNS_172_16_NET is None:
        _MDNS_172_16_NET = ipaddress.IPv4Network("172.16.0.0/12")

    try:
        import psutil
    except Exception:
        psutil = None

    if psutil is not None:
        try:
            if_stats = psutil.net_if_stats()
        except Exception:
            if_stats = {}
        try:
            iter_addrs = list(psutil.net_if_addrs().items())
        except Exception:
            iter_addrs = []

        candidates: list[tuple[int, str, str]] = []  # (priority, nic, ip)
        for nic, addrs in iter_addrs:
            nic_info = if_stats.get(nic)
            if nic_info is not None and not nic_info.isup:
                continue
            for addr in addrs:
                if addr.family != socket.AF_INET:
                    continue
                ip = addr.address
                if not ip or ip.startswith("127.") or ip.startswith("169.254."):
                    continue
                try:
                    ip_obj = ipaddress.IPv4Address(ip)
                except Exception:
                    continue
                # 172.16/12 on Windows is typically Docker/WSL/Hyper-V — skip entirely.
                if ip_obj in _MDNS_172_16_NET:
                    continue
                if not ip_obj.is_private:
                    priority = 4
                elif ip.startswith("192.168."):
                    priority = 0
                elif ip.startswith("10."):
                    priority = 1
                else:
                    priority = 3
                candidates.append((priority, nic, ip))

        if candidates:
            # Stable secondary sort by NIC name keeps the choice deterministic
            # across runs when two adapters share the same priority tier.
            candidates.sort(key=lambda c: (c[0], c[1]))
            return candidates[0][2]

    # Fallback: UDP connect trick. Reject anything we'd have filtered above.
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            if ip and not ip.startswith("127.") and not ip.startswith("169.254."):
                try:
                    if ipaddress.IPv4Address(ip) in _MDNS_172_16_NET:
                        ip = None
                except Exception:
                    pass
                if ip:
                    return ip
    except Exception:
        pass

    return None


def zeroconf_available() -> tuple[bool, str | None]:
    """Check if zeroconf is importable."""
    if _ZEROCONF_IMPORT_ERROR is None:
        return True, None
    return False, str(_ZEROCONF_IMPORT_ERROR)


def _extract_request_token(headers, query_params) -> str | None:
    auth_header = headers.get("authorization", "")
    if auth_header.lower().startswith("bearer "):
        return auth_header[7:].strip() or None

    for header_name in ("x-monitor-token", "x-api-token", "x-token"):
        header_value = headers.get(header_name, "")
        if header_value:
            return header_value.strip()

    for query_name in ("token", "access_token"):
        query_value = query_params.get(query_name)
        if query_value:
            return query_value.strip()

    return None


def _is_authorized(token: str, headers, query_params) -> bool:
    if not token:
        return True
    return _extract_request_token(headers, query_params) == token


def _parse_interval_ms(raw_value) -> int:
    try:
        interval_ms = int(raw_value)
    except (TypeError, ValueError):
        interval_ms = DEFAULT_WS_INTERVAL_MS
    return max(MIN_WS_INTERVAL_MS, min(MAX_WS_INTERVAL_MS, interval_ms))


def _create_app(server_cfg: dict, ready_event: threading.Event):
    async def homepage(request):
        return JSONResponse(
            {
                "service": "little-helper-monitor",
                "auth_required": bool(server_cfg["token"]),
                "mdns": bool(server_cfg.get("mdns", True)),
                "endpoints": {
                    "health": "/health",
                    "monitor": "/api/monitor",
                    "websocket": "/ws/monitor",
                },
            }
        )

    async def healthcheck(request):
        return JSONResponse(
            {
                "status": "ok",
                "auth_required": bool(server_cfg["token"]),
                "bind": {
                    "host": server_cfg["host"],
                    "port": server_cfg["port"],
                },
            }
        )

    async def monitor_snapshot(request):
        if not _is_authorized(server_cfg["token"], request.headers, request.query_params):
            return JSONResponse({"detail": "Unauthorized"}, status_code=401)
        snapshot_type = request.query_params.get("type", "default")
        return JSONResponse(system_overlay.get_monitor_snapshot(type=snapshot_type))

    async def monitor_websocket(websocket):
        if not _is_authorized(server_cfg["token"], websocket.headers, websocket.query_params):
            await websocket.close(code=4401, reason="Unauthorized")
            return

        await websocket.accept()
        interval_ms = _parse_interval_ms(websocket.query_params.get("interval_ms"))
        snapshot_type = websocket.query_params.get("type", "default")

        try:
            while True:
                await websocket.send_json(
                    {
                        "type": "snapshot",
                        "payload": system_overlay.get_monitor_snapshot(max_age_ms=interval_ms, type=snapshot_type),
                    }
                )
                await asyncio.sleep(interval_ms / 1000.0)
        except WebSocketDisconnect:
            pass
        except Exception as exc:
            log.debug(f"Monitor websocket closed with error: {exc}")

    routes = [
        Route("/", homepage),
        Route("/health", healthcheck),
        Route("/api/monitor", monitor_snapshot),
        WebSocketRoute("/ws/monitor", monitor_websocket),
    ]

    @asynccontextmanager
    async def lifespan(_app):
        ready_event.set()
        log.info(
            "Monitor server listening on %s:%s",
            server_cfg["host"],
            server_cfg["port"],
        )
        try:
            yield
        finally:
            log.info("Monitor server shutdown complete")

    try:
        return Starlette(debug=False, routes=routes, lifespan=lifespan)
    except TypeError:
        app = Starlette(debug=False, routes=routes)

        @app.on_event("startup")
        async def _on_startup():
            ready_event.set()
            log.info(
                "Monitor server listening on %s:%s",
                server_cfg["host"],
                server_cfg["port"],
            )

        @app.on_event("shutdown")
        async def _on_shutdown():
            log.info("Monitor server shutdown complete")

        return app


class MonitorServerController:
    def __init__(self):
        self._lock = threading.Lock()
        self._thread: threading.Thread | None = None
        self._server = None
        self._ready_event = threading.Event()
        self._startup_error = None
        self._server_cfg = normalize_monitor_server_config({})
        self._zeroconf: Zeroconf | None = None
        self._mdns_service: ServiceInfo | None = None
        self._mdns_lock = threading.Lock()
        self._mdns_stop_event = threading.Event()
        self._mdns_thread: threading.Thread | None = None
        self._registered_ip: str | None = None

    def _start_mdns(self, server_cfg: dict) -> None:
        """Start the mDNS watcher; it registers as soon as a usable IP exists.

        Polls the network periodically so registration recovers from boot-time
        situations like "no cable yet" and from later changes (DHCP renewal,
        adapter coming up/down, IP change).
        """
        if not server_cfg.get("mdns", True):
            return
        if Zeroconf is None:
            log.warning("mDNS skipped: zeroconf library not available")
            return

        # Guard: stop any previous watcher before spawning a new one.
        self._stop_mdns()
        self._mdns_stop_event.clear()

        thread = threading.Thread(
            target=self._mdns_watcher_loop,
            args=(server_cfg,),
            daemon=True,
            name="monitor-mdns-watcher",
        )
        self._mdns_thread = thread
        thread.start()

    def _mdns_watcher_loop(self, server_cfg: dict) -> None:
        # Do the first check immediately so registration happens without delay.
        try:
            self._mdns_check_and_register(server_cfg)
        except Exception as exc:
            log.debug(f"mDNS watcher initial check error: {exc}")

        # Then enter the periodic polling loop. wait() returns True when the
        # event is set (-> stop), so we sleep first, then re-check.
        while not self._mdns_stop_event.wait(timeout=MDNS_WATCH_INTERVAL_SEC):
            try:
                self._mdns_check_and_register(server_cfg)
            except Exception as exc:
                log.debug(f"mDNS watcher iteration error: {exc}")

    def _mdns_check_and_register(self, server_cfg: dict) -> None:
        new_ip = _get_local_ip()
        port = server_cfg["port"]
        with self._mdns_lock:
            if new_ip is None:
                # Network unavailable. Drop any stale registration so clients
                # don't keep trying an address that no longer works.
                if self._zeroconf is not None:
                    log.info("mDNS: no usable IP, unregistering until network returns")
                    self._do_unregister_mdns_locked()
                return
            if self._zeroconf is None:
                self._do_register_mdns_locked(server_cfg, new_ip, port)
            elif new_ip != self._registered_ip:
                log.info(
                    "mDNS IP changed: %s -> %s, re-registering",
                    self._registered_ip, new_ip,
                )
                self._do_unregister_mdns_locked()
                self._do_register_mdns_locked(server_cfg, new_ip, port)

    def _do_register_mdns_locked(self, server_cfg: dict, local_ip: str, port: int) -> None:
        try:
            ip_bytes = socket.inet_aton(local_ip)
            properties = {
                "auth_required": str(bool(server_cfg["token"])).lower(),
                "path": "/ws/monitor",
            }
            self._mdns_service = ServiceInfo(
                MDNS_SERVICE_TYPE,
                f"Little-Helper-Monitor.{MDNS_SERVICE_TYPE}",
                addresses=[ip_bytes],
                port=port,
                properties=properties,
            )
            self._zeroconf = Zeroconf()
            self._zeroconf.register_service(self._mdns_service)
            self._registered_ip = local_ip
            log.info(
                "mDNS service registered: Little-Helper-Monitor.%s at %s:%s",
                MDNS_SERVICE_TYPE, local_ip, port,
            )
        except Exception as exc:
            log.warning(f"mDNS registration failed: {exc}")
            self._do_unregister_mdns_locked()

    def _do_unregister_mdns_locked(self) -> None:
        if self._zeroconf is not None:
            try:
                if self._mdns_service is not None:
                    self._zeroconf.unregister_service(self._mdns_service)
                self._zeroconf.close()
            except Exception as exc:
                log.debug(f"mDNS cleanup error: {exc}")
            finally:
                self._zeroconf = None
                self._mdns_service = None
                self._registered_ip = None

    def _stop_mdns(self) -> None:
        """Stop the watcher and unregister the mDNS service."""
        self._mdns_stop_event.set()
        thread = self._mdns_thread
        self._mdns_thread = None
        # Don't join from inside the watcher thread itself.
        if thread is not None and thread is not threading.current_thread() and thread.is_alive():
            thread.join(timeout=1)
        with self._mdns_lock:
            self._do_unregister_mdns_locked()

    def start(self, config: dict) -> bool:
        server_cfg = normalize_monitor_server_config(config)
        if not server_cfg["enabled"]:
            self.stop()
            return False

        deps_ok, deps_error = monitor_server_dependencies_available()
        if not deps_ok:
            raise RuntimeError(f"Monitor server dependencies are unavailable: {deps_error}")

        with self._lock:
            if self.is_running() and self._server_cfg == server_cfg:
                return True

        self.stop()

        self._ready_event.clear()
        self._startup_error = None
        self._server_cfg = server_cfg
        thread = threading.Thread(
            target=self._run_server,
            args=(server_cfg,),
            daemon=True,
            name="monitor-server",
        )

        with self._lock:
            self._thread = thread

        thread.start()
        if not self._ready_event.wait(timeout=3):
            if self._startup_error is not None:
                raise RuntimeError(self._startup_error)
            if not thread.is_alive():
                raise RuntimeError(
                    f"Monitor server failed to start on {server_cfg['host']}:{server_cfg['port']}"
                )
            raise RuntimeError("Monitor server startup timed out")
        if self._startup_error is not None:
            raise RuntimeError(self._startup_error)
        if not thread.is_alive():
            raise RuntimeError(
                f"Monitor server failed to start on {server_cfg['host']}:{server_cfg['port']}"
            )
        # mDNS is now started inside _run_server (non-blocking)
        return True

    def stop(self) -> None:
        self._stop_mdns()
        thread = None
        server = None
        with self._lock:
            thread = self._thread
            server = self._server
            self._thread = None
            self._server = None
        if server is not None:
            server.should_exit = True
        if thread is not None and thread.is_alive():
            thread.join(timeout=2)
            if thread.is_alive() and server is not None:
                server.force_exit = True
                thread.join(timeout=0.5)

    def restart(self, config: dict) -> bool:
        self.stop()
        return self.start(config)

    def is_running(self) -> bool:
        return self._thread is not None and self._thread.is_alive()

    def current_config(self) -> dict:
        return dict(self._server_cfg)

    def _run_server(self, server_cfg: dict) -> None:
        try:
            app = _create_app(server_cfg, self._ready_event)
            config = uvicorn.Config(
                app,
                host=server_cfg["host"],
                port=server_cfg["port"],
                log_level="warning",
                log_config=None,
                access_log=False,
                server_header=False,
                ws="websockets",
                lifespan="on",
            )
            server = uvicorn.Server(config)
            with self._lock:
                self._server = server
            # Start mDNS in a daemon thread so it doesn't block server startup.
            self._start_mdns(server_cfg)
            server.run()
            if not server.started and self._startup_error is None:
                self._startup_error = (
                    f"Monitor server failed to bind {server_cfg['host']}:{server_cfg['port']}"
                )
                self._ready_event.set()
        except Exception as exc:
            log.error(f"Monitor server crashed: {exc}", exc_info=True)
            self._startup_error = str(exc)
            self._ready_event.set()
        finally:
            self._stop_mdns()
            with self._lock:
                self._server = None
                if self._thread is not None and not self._thread.is_alive():
                    self._thread = None


_controller = MonitorServerController()


def start_monitor_server(config: dict) -> bool:
    return _controller.start(config)


def stop_monitor_server() -> None:
    _controller.stop()


def restart_monitor_server(config: dict) -> bool:
    return _controller.restart(config)


def monitor_server_is_running() -> bool:
    return _controller.is_running()


def get_running_monitor_server_config() -> dict:
    return _controller.current_config()