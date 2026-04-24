"""
Test mDNS discovery for Little Helper Monitor Server.

Listens for the _little-helper-monitor._tcp.local. service on the local
network and prints discovered instances. Exits after a timeout or when
at least one service is found.

Usage:
    python tests/test_mdns_discovery.py [--timeout 10]
"""

import argparse
import socket
import sys
import time

try:
    from zeroconf import Zeroconf, ServiceBrowser, ServiceStateChange
except ImportError:
    print("ERROR: zeroconf is not installed. Run: pip install zeroconf")
    sys.exit(1)

MDNS_SERVICE_TYPE = "_lhm._tcp.local."


def resolve_hostname(name: str) -> str | None:
    """Try to resolve a .local mDNS hostname to an IP address."""
    try:
        return socket.gethostbyname(name)
    except socket.gaierror:
        return None


def main():
    parser = argparse.ArgumentParser(description="Discover Little Helper Monitor Server via mDNS")
    parser.add_argument("--timeout", type=float, default=10, help="Discovery timeout in seconds (default: 10)")
    args = parser.parse_args()

    discovered = []
    done_event = __import__("threading").Event()

    def on_service_state_change(zeroconf: Zeroconf, service_type: str, name: str, state_change: ServiceStateChange):
        if state_change == ServiceStateChange.Added:
            info = zeroconf.get_service_info(service_type, name)
            if info is None:
                return

            # Extract IP addresses
            addresses = []
            for addr in info.addresses:
                try:
                    addresses.append(socket.inet_ntoa(addr))
                except Exception:
                    pass

            port = info.port
            properties = {}
            if info.properties:
                for key, val in info.properties.items():
                    if isinstance(key, bytes):
                        key = key.decode("utf-8", errors="replace")
                    if isinstance(val, bytes):
                        val = val.decode("utf-8", errors="replace")
                    properties[key] = val

            server_name = info.server.rstrip(".") if info.server else "unknown"

            print("=" * 60)
            print(f"  Found: {name}")
            print(f"  Host:  {server_name}")
            print(f"  Addrs: {', '.join(addresses) if addresses else 'N/A'}")
            print(f"  Port:  {port}")
            print(f"  Props: {properties}")
            print(f"  WS URL: ws://{addresses[0] if addresses else server_name}:{port}{properties.get('path', '/ws/monitor')}")
            print("=" * 60)

            discovered.append({
                "name": name,
                "host": server_name,
                "addresses": addresses,
                "port": port,
                "properties": properties,
            })
            done_event.set()

    print(f"Listening for mDNS service: {MDNS_SERVICE_TYPE}")
    print(f"Timeout: {args.timeout}s\n")

    zeroconf = Zeroconf()
    browser = ServiceBrowser(zeroconf, MDNS_SERVICE_TYPE, handlers=[on_service_state_change])

    try:
        done_event.wait(timeout=args.timeout)
    finally:
        browser.cancel()
        zeroconf.close()

    if discovered:
        print(f"\n✓ Discovered {len(discovered)} service(s)")
    else:
        print(f"\n✗ No services found within {args.timeout}s")
        print("  Make sure the Monitor Server is running with mDNS enabled.")
        sys.exit(1)


if __name__ == "__main__":
    main()
