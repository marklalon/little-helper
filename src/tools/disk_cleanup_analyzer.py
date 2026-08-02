"""Windows disk usage analyzer with explicitly confirmed permanent cleanup.

Scanning is read-only. The tool measures the drive and conservatively estimates
reclaimable space; it permanently deletes old temporary files and expired
pip/npm caches only after the user enters the exact confirmation token ``CLEAN``.

Examples::

    python src/tools/disk_cleanup_analyzer.py
    python src/tools/disk_cleanup_analyzer.py --full
    python src/tools/disk_cleanup_analyzer.py --scan-only
    python src/tools/disk_cleanup_analyzer.py --json > disk-report.json
"""

from __future__ import annotations

import argparse
import heapq
import json
import os
import shutil
import stat
import sys
import threading
import time
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, as_completed, wait
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Sequence


REPARSE_POINT_ATTRIBUTE = 0x400
LEVEL_RECOMMENDED = "recommended"
LEVEL_REVIEW = "review"
LEVEL_SYSTEM = "system"


@dataclass(frozen=True)
class CandidateSpec:
    name: str
    path: Path
    level: str
    description: str
    min_age_days: Optional[int] = None
    reclaimable: bool = True


@dataclass
class ScanResult:
    path: str
    total_bytes: int = 0
    reclaimable_bytes: int = 0
    file_count: int = 0
    reclaimable_file_count: int = 0
    error_count: int = 0
    largest_files: Optional[list[tuple[int, str]]] = None


@dataclass
class CleanupResult:
    deleted_bytes: int = 0
    deleted_file_count: int = 0
    removed_directory_count: int = 0
    error_count: int = 0
    failed_paths: Optional[list[str]] = None


class ScanProgress:
    """Thread-safe, periodically refreshed CLI scan progress."""

    def __init__(
        self,
        *,
        enabled: bool = True,
        stream=None,
        refresh_interval: float = 1.0,
    ) -> None:
        self.enabled = enabled
        self.stream = stream if stream is not None else sys.stderr
        self.refresh_interval = max(0.1, refresh_interval)
        self._lock = threading.Lock()
        self._stop_event = threading.Event()
        self._thread: Optional[threading.Thread] = None
        self._phase = "准备扫描"
        self._action = "已扫描"
        self._phase_started = time.monotonic()
        self._total_items = 0
        self._completed_items = 0
        self._files = 0
        self._bytes = 0
        self._errors = 0
        self._active: set[str] = set()
        self._last_line_length = 0
        self._line_open = False

    def __enter__(self) -> "ScanProgress":
        if self.enabled:
            self._thread = threading.Thread(
                target=self._refresh_loop,
                name="disk-scan-progress",
                daemon=True,
            )
            self._thread.start()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self._stop_event.set()
        if self._thread is not None:
            self._thread.join(timeout=2)
        if self.enabled and self._line_open:
            print(file=self.stream, flush=True)
            self._line_open = False

    def begin_phase(self, label: str, total_items: int, action: str = "已扫描") -> None:
        if not self.enabled:
            return
        with self._lock:
            if self._line_open:
                print(file=self.stream, flush=True)
                self._line_open = False
                self._last_line_length = 0
            self._phase = label
            self._action = action
            self._phase_started = time.monotonic()
            self._total_items = total_items
            self._completed_items = 0
            self._files = 0
            self._bytes = 0
            self._errors = 0
            self._active.clear()
        self._render()

    def item_started(self, path: Path) -> None:
        if self.enabled:
            with self._lock:
                self._active.add(str(path))

    def advance(self, files: int, size_bytes: int, errors: int = 0) -> None:
        if self.enabled:
            with self._lock:
                self._files += files
                self._bytes += size_bytes
                self._errors += errors

    def item_finished(self, path: Path) -> None:
        if self.enabled:
            with self._lock:
                self._active.discard(str(path))
                self._completed_items += 1

    def finish_phase(self) -> None:
        if not self.enabled:
            return
        self._render(final=True)

    def _refresh_loop(self) -> None:
        while not self._stop_event.wait(self.refresh_interval):
            self._render()

    def _render(self, final: bool = False) -> None:
        if not self.enabled:
            return
        with self._lock:
            elapsed = max(0.0, time.monotonic() - self._phase_started)
            item_progress = f"{self._completed_items}/{self._total_items} 项"
            active_names = sorted(Path(path).name or path for path in self._active)
            active = "、".join(active_names[:2])
            if len(active_names) > 2:
                active += f" 等 {len(active_names)} 项"
            line = (
                f"[{self._phase}] {item_progress} | {self._action} {self._files:,} 个文件 / "
                f"{format_bytes(self._bytes)} | {elapsed:.1f} 秒"
            )
            if self._errors:
                line += f" | 无法访问 {self._errors} 处"
            if active and not final:
                line += f" | 正在处理：{active}"

            is_tty = bool(getattr(self.stream, "isatty", lambda: False)())
            if is_tty and not final:
                padded = line.ljust(self._last_line_length)
                print(f"\r{padded}", end="", file=self.stream, flush=True)
                self._last_line_length = len(line)
                self._line_open = True
            else:
                if is_tty and self._line_open:
                    padded = line.ljust(self._last_line_length)
                    print(f"\r{padded}", file=self.stream, flush=True)
                else:
                    print(line, file=self.stream, flush=True)
                self._last_line_length = 0
                self._line_open = False


def format_bytes(value: int) -> str:
    """Format a byte count using binary units."""
    size = float(max(0, value))
    units = ("B", "KiB", "MiB", "GiB", "TiB", "PiB")
    for unit in units:
        if size < 1024 or unit == units[-1]:
            return f"{size:.0f} {unit}" if unit == "B" else f"{size:.2f} {unit}"
        size /= 1024
    return f"{size:.2f} PiB"


def _is_reparse_point(stat_result: os.stat_result) -> bool:
    attributes = getattr(stat_result, "st_file_attributes", 0)
    return bool(attributes & REPARSE_POINT_ATTRIBUTE)


def _push_largest(heap: list[tuple[int, str]], limit: int, size: int, path: str) -> None:
    if limit <= 0:
        return
    item = (size, path)
    if len(heap) < limit:
        heapq.heappush(heap, item)
    elif size > heap[0][0]:
        heapq.heapreplace(heap, item)


def scan_path(
    path: Path,
    *,
    cutoff_timestamp: Optional[float] = None,
    all_reclaimable: bool = False,
    largest_limit: int = 0,
    progress: Optional[ScanProgress] = None,
) -> ScanResult:
    """Scan one file or tree without following symlinks or junctions.

    Sizes are logical file sizes, so compressed, sparse, and hard-linked files
    can make the result differ from the exact number of allocated disk bytes.
    """
    result = ScanResult(path=str(path), largest_files=[])
    pending = [path]
    largest: list[tuple[int, str]] = []
    pending_files = 0
    pending_bytes = 0
    pending_errors = 0

    def flush_progress() -> None:
        nonlocal pending_files, pending_bytes, pending_errors
        if progress is not None and (pending_files or pending_bytes or pending_errors):
            progress.advance(pending_files, pending_bytes, pending_errors)
        pending_files = 0
        pending_bytes = 0
        pending_errors = 0

    while pending:
        current = pending.pop()
        try:
            stat_result = os.lstat(current)
            if current.is_symlink() or _is_reparse_point(stat_result):
                continue
            if stat.S_ISDIR(stat_result.st_mode):
                try:
                    with os.scandir(current) as entries:
                        pending.extend(Path(entry.path) for entry in entries)
                except (OSError, PermissionError):
                    result.error_count += 1
                    pending_errors += 1
                    flush_progress()
                continue

            size = stat_result.st_size
            result.total_bytes += size
            result.file_count += 1
            pending_files += 1
            pending_bytes += size
            eligible = all_reclaimable or (
                cutoff_timestamp is not None and stat_result.st_mtime <= cutoff_timestamp
            )
            if eligible:
                result.reclaimable_bytes += size
                result.reclaimable_file_count += 1
            _push_largest(largest, largest_limit, size, str(current))
            if pending_files >= 256:
                flush_progress()
        except (OSError, PermissionError):
            result.error_count += 1
            pending_errors += 1
            flush_progress()

    flush_progress()
    result.largest_files = sorted(largest, reverse=True)
    return result


def _profile_directories(root: Path) -> list[Path]:
    users_root = root / "Users"
    try:
        return [
            Path(entry.path)
            for entry in os.scandir(users_root)
            if entry.is_dir(follow_symlinks=False)
            and entry.name.lower() not in {"all users", "default", "default user", "public"}
        ]
    except OSError:
        return []


def build_candidate_specs(
    root: Path, temp_age_days: int = 7, cache_age_days: int = 30
) -> list[CandidateSpec]:
    """Build conservative cleanup candidates for a Windows system drive."""
    specs: list[CandidateSpec] = [
        CandidateSpec(
            "Windows 旧临时文件",
            root / "Windows" / "Temp",
            LEVEL_RECOMMENDED,
            "超过保留天数的系统临时文件；正在使用或无权限的文件不会被本工具处理",
            temp_age_days,
        ),
        CandidateSpec(
            "回收站",
            root / "$Recycle.Bin",
            LEVEL_REVIEW,
            "清空后难以恢复，请先检查回收站内容",
            None,
        ),
        CandidateSpec(
            "Windows 错误报告",
            root / "ProgramData" / "Microsoft" / "Windows" / "WER" / "ReportArchive",
            LEVEL_REVIEW,
            "仅在不需要排查历史崩溃时考虑清理",
            cache_age_days,
        ),
        CandidateSpec(
            "排队中的 Windows 错误报告",
            root / "ProgramData" / "Microsoft" / "Windows" / "WER" / "ReportQueue",
            LEVEL_REVIEW,
            "可能尚未上传；仅在不需要故障诊断时考虑清理",
            cache_age_days,
        ),
        CandidateSpec(
            "Windows 小型内存转储",
            root / "Windows" / "Minidump",
            LEVEL_REVIEW,
            "删除后无法再用这些文件分析蓝屏原因",
            cache_age_days,
        ),
        CandidateSpec(
            "完整内存转储",
            root / "Windows" / "MEMORY.DMP",
            LEVEL_REVIEW,
            "删除后无法再用此文件分析蓝屏原因",
            cache_age_days,
        ),
        CandidateSpec(
            "Windows.old",
            root / "Windows.old",
            LEVEL_SYSTEM,
            "请使用“设置 > 系统 > 存储 > 临时文件”删除；删除后不能回退旧版本",
            None,
        ),
        CandidateSpec(
            "传递优化缓存",
            root / "ProgramData" / "Microsoft" / "Windows" / "DeliveryOptimization" / "Cache",
            LEVEL_SYSTEM,
            "请通过 Windows 存储设置或磁盘清理处理",
            None,
        ),
        CandidateSpec(
            "Windows 更新下载缓存（占用参考）",
            root / "Windows" / "SoftwareDistribution" / "Download",
            LEVEL_SYSTEM,
            "不要手动删除；显示的是占用量，不等于 Windows 实际可清理量",
            None,
            reclaimable=False,
        ),
    ]

    for profile in _profile_directories(root):
        profile_name = profile.name
        local = profile / "AppData" / "Local"
        specs.extend(
            [
                CandidateSpec(
                    f"{profile_name} 的旧临时文件",
                    local / "Temp",
                    LEVEL_RECOMMENDED,
                    "超过保留天数的用户临时文件",
                    temp_age_days,
                ),
                CandidateSpec(
                    f"{profile_name} 的 DirectX 着色器缓存",
                    local / "D3DSCache",
                    LEVEL_REVIEW,
                    "可自动重建，但首次启动程序时可能变慢",
                    cache_age_days,
                ),
                CandidateSpec(
                    f"{profile_name} 的应用崩溃转储",
                    local / "CrashDumps",
                    LEVEL_REVIEW,
                    "删除后无法用于排查历史应用崩溃",
                    cache_age_days,
                ),
                CandidateSpec(
                    f"{profile_name} 的 pip 缓存",
                    local / "pip" / "Cache",
                    LEVEL_RECOMMENDED,
                    "超过保留天数后可清理；可重新下载，不影响已安装的 Python 包",
                    cache_age_days,
                ),
                CandidateSpec(
                    f"{profile_name} 的 npm 缓存",
                    local / "npm-cache",
                    LEVEL_RECOMMENDED,
                    "超过保留天数后可清理；可重新下载，不影响已安装的 Node.js 包",
                    cache_age_days,
                ),
            ]
        )

    # A normalized-path pass prevents accidental double counting if profiles or
    # future candidate rules resolve to the same location.
    unique: list[CandidateSpec] = []
    seen: set[str] = set()
    for spec in specs:
        key = os.path.normcase(os.path.abspath(spec.path))
        if key not in seen:
            seen.add(key)
            unique.append(spec)
    return unique


def _scan_candidate(
    spec: CandidateSpec, now: float, progress: Optional[ScanProgress] = None
) -> dict:
    if progress is not None:
        progress.item_started(spec.path)
    cutoff = None
    if spec.min_age_days is not None:
        cutoff = now - spec.min_age_days * 86400
    result = scan_path(
        spec.path,
        cutoff_timestamp=cutoff,
        all_reclaimable=spec.min_age_days is None,
        progress=progress,
    )
    if progress is not None:
        progress.item_finished(spec.path)
    reclaimable_bytes = result.reclaimable_bytes if spec.reclaimable else 0
    return {
        "name": spec.name,
        "path": str(spec.path),
        "level": spec.level,
        "description": spec.description,
        "min_age_days": spec.min_age_days,
        "occupied_bytes": result.total_bytes,
        "reclaimable_bytes": reclaimable_bytes,
        "file_count": result.file_count,
        "reclaimable_file_count": result.reclaimable_file_count if spec.reclaimable else 0,
        "error_count": result.error_count,
        "reclaimable": spec.reclaimable,
    }


def scan_candidates(
    specs: Iterable[CandidateSpec],
    workers: int,
    now: float,
    progress: Optional[ScanProgress] = None,
) -> list[dict]:
    existing = [spec for spec in specs if spec.path.exists()]
    if progress is not None:
        progress.begin_phase("扫描清理候选", len(existing))
    if not existing:
        if progress is not None:
            progress.finish_phase()
        return []
    results: list[dict] = []
    with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
        futures = {
            executor.submit(_scan_candidate, spec, now, progress): spec for spec in existing
        }
        for future in as_completed(futures):
            results.append(future.result())
    if progress is not None:
        progress.finish_phase()
    return sorted(results, key=lambda item: item["reclaimable_bytes"], reverse=True)


def _scan_overview_item(
    path: Path, largest_limit: int, progress: Optional[ScanProgress]
) -> ScanResult:
    if progress is not None:
        progress.item_started(path)
    try:
        return scan_path(path, largest_limit=largest_limit, progress=progress)
    finally:
        if progress is not None:
            progress.item_finished(path)


def scan_overview(
    root: Path,
    workers: int,
    largest_limit: int,
    progress: Optional[ScanProgress] = None,
) -> tuple[list[dict], list[dict], int]:
    """Scan each top-level entry and return sizes plus the largest files."""
    items: list[tuple[Path, bool]] = []
    root_errors = 0
    try:
        with os.scandir(root) as entries:
            for entry in entries:
                try:
                    stat_result = entry.stat(follow_symlinks=False)
                    if entry.is_symlink() or _is_reparse_point(stat_result):
                        continue
                    items.append((Path(entry.path), entry.is_dir(follow_symlinks=False)))
                except OSError:
                    root_errors += 1
    except OSError:
        return [], [], 1

    if progress is not None:
        progress.begin_phase("统计全盘占用", len(items))
    results: list[ScanResult] = []
    with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
        futures = {
            executor.submit(_scan_overview_item, path, largest_limit, progress): (path, is_dir)
            for path, is_dir in items
        }
        for future in as_completed(futures):
            results.append(future.result())

    if progress is not None:
        progress.finish_phase()

    overview = [
        {
            "path": result.path,
            "size_bytes": result.total_bytes,
            "file_count": result.file_count,
            "error_count": result.error_count,
        }
        for result in sorted(results, key=lambda value: value.total_bytes, reverse=True)
    ]
    largest_heap: list[tuple[int, str]] = []
    for result in results:
        for size, path in result.largest_files or []:
            _push_largest(largest_heap, largest_limit, size, path)
    largest = [
        {"path": path, "size_bytes": size}
        for size, path in sorted(largest_heap, reverse=True)
    ]
    return overview, largest, root_errors + sum(item.error_count for item in results)


def _merge_cleanup_result(target: CleanupResult, source: CleanupResult) -> None:
    target.deleted_bytes += source.deleted_bytes
    target.deleted_file_count += source.deleted_file_count
    target.removed_directory_count += source.removed_directory_count
    target.error_count += source.error_count
    if source.failed_paths:
        if target.failed_paths is None:
            target.failed_paths = []
        remaining = max(0, 20 - len(target.failed_paths))
        target.failed_paths.extend(source.failed_paths[:remaining])


def _delete_file_batch(
    files: list[tuple[Path, int]], progress: Optional[ScanProgress]
) -> CleanupResult:
    result = CleanupResult(failed_paths=[])
    for path, size in files:
        try:
            os.unlink(path)
        except PermissionError:
            # Read-only temp files are still within the explicitly confirmed
            # cleanup roots. Clear only the write-protection bit and retry once.
            try:
                os.chmod(path, stat.S_IWRITE)
                os.unlink(path)
            except OSError:
                result.error_count += 1
                if len(result.failed_paths) < 20:
                    result.failed_paths.append(str(path))
                if progress is not None:
                    progress.advance(0, 0, 1)
                continue
        except OSError:
            result.error_count += 1
            if len(result.failed_paths) < 20:
                result.failed_paths.append(str(path))
            if progress is not None:
                progress.advance(0, 0, 1)
            continue

        result.deleted_bytes += size
        result.deleted_file_count += 1
        if progress is not None:
            progress.advance(1, size)
    return result


def _validated_cleanup_specs(root: Path, specs: Iterable[CandidateSpec]) -> list[CandidateSpec]:
    """Keep only recommended candidates strictly below the analyzed root."""
    root_resolved = root.resolve()
    validated: list[CandidateSpec] = []
    for spec in specs:
        if spec.level != LEVEL_RECOMMENDED or not spec.reclaimable:
            continue
        try:
            candidate = spec.path.resolve(strict=False)
            candidate.relative_to(root_resolved)
        except (OSError, ValueError):
            continue
        if candidate == root_resolved:
            continue
        validated.append(spec)
    return validated


def clean_recommended_files(
    root: Path,
    specs: Iterable[CandidateSpec],
    *,
    workers: int = 4,
    now: Optional[float] = None,
    progress: Optional[ScanProgress] = None,
) -> CleanupResult:
    """Permanently delete eligible files from recommended cleanup roots only."""
    roots = [spec for spec in _validated_cleanup_specs(root, specs) if spec.path.exists()]
    result = CleanupResult(failed_paths=[])
    if progress is not None:
        progress.begin_phase("永久清理临时文件和开发缓存", len(roots), action="已删除")
    if not roots:
        if progress is not None:
            progress.finish_phase()
        return result

    scan_time = time.time() if now is None else now
    directories: list[Path] = []
    batch: list[tuple[Path, int]] = []
    pending_futures = set()
    max_pending = max(1, workers) * 2

    def collect_completed(done_futures) -> None:
        for future in done_futures:
            _merge_cleanup_result(result, future.result())

    with ThreadPoolExecutor(max_workers=max(1, workers)) as executor:
        for spec in roots:
            if progress is not None:
                progress.item_started(spec.path)
            cutoff = None
            if spec.min_age_days is not None:
                cutoff = scan_time - spec.min_age_days * 86400
            pending_paths = [spec.path]

            while pending_paths:
                current = pending_paths.pop()
                try:
                    stat_result = os.lstat(current)
                    if current.is_symlink() or _is_reparse_point(stat_result):
                        continue
                    if stat.S_ISDIR(stat_result.st_mode):
                        if current != spec.path:
                            directories.append(current)
                        try:
                            with os.scandir(current) as entries:
                                pending_paths.extend(Path(entry.path) for entry in entries)
                        except OSError:
                            result.error_count += 1
                            if len(result.failed_paths) < 20:
                                result.failed_paths.append(str(current))
                            if progress is not None:
                                progress.advance(0, 0, 1)
                        continue

                    if cutoff is not None and stat_result.st_mtime > cutoff:
                        continue
                    batch.append((current, stat_result.st_size))
                    if len(batch) >= 256:
                        pending_futures.add(
                            executor.submit(_delete_file_batch, batch, progress)
                        )
                        batch = []
                        if len(pending_futures) >= max_pending:
                            done, pending_futures = wait(
                                pending_futures, return_when=FIRST_COMPLETED
                            )
                            collect_completed(done)
                except OSError:
                    result.error_count += 1
                    if len(result.failed_paths) < 20:
                        result.failed_paths.append(str(current))
                    if progress is not None:
                        progress.advance(0, 0, 1)

            if progress is not None:
                progress.item_finished(spec.path)

        if batch:
            pending_futures.add(executor.submit(_delete_file_batch, batch, progress))
        if pending_futures:
            done, _ = wait(pending_futures)
            collect_completed(done)

    # Remove only now-empty subdirectories. Candidate roots themselves remain.
    for directory in sorted(directories, key=lambda path: len(path.parts), reverse=True):
        try:
            os.rmdir(directory)
            result.removed_directory_count += 1
        except OSError:
            pass

    if progress is not None:
        progress.finish_phase()
    return result


def analyze_drive(
    root: Path,
    *,
    temp_age_days: int = 7,
    cache_age_days: int = 30,
    workers: int = 4,
    full_scan: bool = True,
    largest_limit: int = 20,
    now: Optional[float] = None,
    progress: Optional[ScanProgress] = None,
) -> dict:
    """Analyze a drive or test root and return a JSON-serializable report."""
    root = Path(os.path.abspath(root))
    usage = shutil.disk_usage(root)
    specs = build_candidate_specs(root, temp_age_days, cache_age_days)
    candidates = scan_candidates(
        specs,
        workers,
        now if now is not None else time.time(),
        progress,
    )

    overview: list[dict] = []
    largest_files: list[dict] = []
    overview_errors = 0
    if full_scan:
        overview, largest_files, overview_errors = scan_overview(
            root, workers, largest_limit, progress
        )

    reclaimable = {
        level: sum(
            item["reclaimable_bytes"] for item in candidates if item["level"] == level
        )
        for level in (LEVEL_RECOMMENDED, LEVEL_REVIEW, LEVEL_SYSTEM)
    }
    protected_system_files = []
    for filename, description in (
        ("hiberfil.sys", "休眠文件；只能通过更改休眠设置释放，会影响休眠/快速启动"),
        ("pagefile.sys", "分页文件；不要作为普通垃圾文件删除"),
        ("swapfile.sys", "系统交换文件；不要手动删除"),
    ):
        path = root / filename
        try:
            size = path.stat().st_size
        except OSError:
            continue
        protected_system_files.append(
            {"path": str(path), "size_bytes": size, "description": description}
        )

    return {
        "schema_version": 1,
        "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
        "root": str(root),
        "disk": {"total_bytes": usage.total, "used_bytes": usage.used, "free_bytes": usage.free},
        "policy": {
            "temp_age_days": temp_age_days,
            "cache_age_days": cache_age_days,
            "scan_read_only": True,
            "cleanup_requires_confirmation": "CLEAN",
            "size_basis": "logical_file_size_estimate",
        },
        "reclaimable_summary": reclaimable,
        "candidates": candidates,
        "protected_system_files": protected_system_files,
        "overview": overview,
        "largest_files": largest_files,
        "overview_error_count": overview_errors,
        "full_scan": full_scan,
    }


def _print_report(report: dict, top: int) -> None:
    disk = report["disk"]
    used_percent = disk["used_bytes"] / disk["total_bytes"] * 100 if disk["total_bytes"] else 0
    summary = report["reclaimable_summary"]
    print(f"磁盘空间分析（只读）：{report['root']}")
    print(
        f"容量 {format_bytes(disk['total_bytes'])} | 已用 {format_bytes(disk['used_bytes'])} "
        f"({used_percent:.1f}%) | 可用 {format_bytes(disk['free_bytes'])}"
    )
    print(
        f"\n较安全的清理估算：{format_bytes(summary[LEVEL_RECOMMENDED])}"
        f"（临时文件 ≥{report['policy']['temp_age_days']} 天；"
        f"pip/npm 缓存 ≥{report['policy']['cache_age_days']} 天）"
    )
    print(
        f"确认后可额外释放：{format_bytes(summary[LEVEL_REVIEW])} | "
        f"需由 Windows 处理：{format_bytes(summary[LEVEL_SYSTEM])}"
    )

    labels = {
        LEVEL_RECOMMENDED: "较安全",
        LEVEL_REVIEW: "需确认",
        LEVEL_SYSTEM: "系统处理",
    }
    visible = [
        item for item in report["candidates"]
        if item["occupied_bytes"] > 0 or item["error_count"] > 0
    ]
    if visible:
        print("\n清理候选：")
        for item in visible:
            size = item["reclaimable_bytes"] if item["reclaimable"] else item["occupied_bytes"]
            suffix = "可释放估算" if item["reclaimable"] else "当前占用（不计入可释放）"
            age = f"，仅计 ≥{item['min_age_days']} 天" if item["min_age_days"] is not None else ""
            errors = f"，{item['error_count']} 处无法访问" if item["error_count"] else ""
            print(f"  [{labels[item['level']]}] {item['name']}: {format_bytes(size)} {suffix}{age}{errors}")
            print(f"      {item['path']}")
            print(f"      {item['description']}")

    if report["protected_system_files"]:
        print("\n受保护/特殊系统文件（不计入清理空间）：")
        for item in report["protected_system_files"]:
            print(f"  {format_bytes(item['size_bytes']):>12}  {item['path']}")
            print(f"                {item['description']}")

    if report["full_scan"]:
        print(f"\n顶层占用（前 {top} 项）：")
        for item in report["overview"][:top]:
            errors = f"  [无法访问 {item['error_count']} 处]" if item["error_count"] else ""
            print(f"  {format_bytes(item['size_bytes']):>12}  {item['path']}{errors}")
        if report["largest_files"]:
            print(f"\n最大文件（前 {top} 项，仅供人工检查，不代表可删除）：")
            for item in report["largest_files"][:top]:
                print(f"  {format_bytes(item['size_bytes']):>12}  {item['path']}")
    else:
        print("\n当前为默认快速模式，未遍历整个磁盘；添加 --full 可查看顶层占用和最大文件。")

    total_errors = report["overview_error_count"] + sum(
        item["error_count"] for item in report["candidates"]
    )
    print(
        "\n说明：扫描阶段只读；只有随后输入 CLEAN 才会永久删除“较安全”的旧临时文件。"
        "结果为逻辑文件大小估算，实际释放量可能不同。"
    )
    if total_errors:
        print(f"有 {total_errors} 处无法访问，报告可能低估占用；以管理员身份运行可提高完整度。")


def _prompt_for_cleanup(report: dict, stream=None, input_stream=None) -> bool:
    """Ask for the exact destructive confirmation token."""
    output = stream if stream is not None else sys.stdout
    source = input_stream if input_stream is not None else sys.stdin
    cleanable = report["reclaimable_summary"][LEVEL_RECOMMENDED]
    if cleanable <= 0:
        print("\n没有发现符合保留天数条件的较安全清理文件。", file=output)
        return False

    print(
        "\n准备永久清理以下“较安全”旧临时文件和 pip/npm 缓存（不会进入回收站）：",
        file=output,
    )
    for item in report["candidates"]:
        if item["level"] == LEVEL_RECOMMENDED and item["reclaimable_bytes"] > 0:
            print(
                f"  {item['name']}: {format_bytes(item['reclaimable_bytes'])}，"
                f"{item['reclaimable_file_count']:,} 个文件",
                file=output,
            )
    print(f"预计最多永久释放：{format_bytes(cleanable)}", file=output)
    print("输入 CLEAN 确认永久删除，输入其他内容取消：", end="", file=output, flush=True)
    response = source.readline()
    return response.strip().casefold() == "clean"


def _prompt_for_direct_cleanup(
    temp_age_days: int,
    cache_age_days: int,
    stream=None,
    input_stream=None,
) -> bool:
    """Confirm cleanup when the user deliberately skipped the estimate pass."""
    output = stream if stream is not None else sys.stdout
    source = input_stream if input_stream is not None else sys.stdin
    print("\n已跳过空间分析，将直接进行以下永久清理（不会进入回收站）：", file=output)
    print(f"  超过 {temp_age_days} 天的 Windows/用户临时文件", file=output)
    print(f"  超过 {cache_age_days} 天的 pip/npm 缓存", file=output)
    print("不会清理回收站、崩溃转储、Windows 更新文件或其他缓存。", file=output)
    print("输入 CLEAN 确认永久删除，输入其他内容取消：", end="", file=output, flush=True)
    response = source.readline()
    return response.strip().casefold() == "clean"


def _cleanup_result_dict(result: CleanupResult, confirmed: bool) -> dict:
    return {
        "confirmed": confirmed,
        "permanent_delete": True,
        "deleted_bytes": result.deleted_bytes,
        "deleted_file_count": result.deleted_file_count,
        "removed_directory_count": result.removed_directory_count,
        "error_count": result.error_count,
        "failed_paths": result.failed_paths or [],
    }


def _print_cleanup_result(result: CleanupResult) -> None:
    print(
        f"\n清理完成：永久删除 {result.deleted_file_count:,} 个文件，"
        f"实际释放约 {format_bytes(result.deleted_bytes)}，"
        f"移除 {result.removed_directory_count:,} 个空目录。"
    )
    if result.error_count:
        print(f"有 {result.error_count:,} 个文件或目录清理失败（通常是正在使用或权限不足）：")
        for path in result.failed_paths or []:
            print(f"  {path}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="分析 Windows 磁盘占用，并在明确确认后永久清理旧临时文件和 pip/npm 缓存。"
    )
    parser.add_argument("--drive", default="C:\\", help="要分析的盘符或目录（默认：C:\\）")
    parser.add_argument("--full", action="store_true", help="完整遍历磁盘，统计顶层占用和最大文件")
    parser.add_argument(
        "--skip-analyze",
        action="store_true",
        help="跳过空间分析，确认后直接筛选并永久清理",
    )
    parser.add_argument("--json", action="store_true", help="输出机器可读的 JSON")
    parser.add_argument("--no-progress", action="store_true", help="不显示扫描进度")
    parser.add_argument("--scan-only", action="store_true", help="只生成报告，不询问或执行清理")
    parser.add_argument("--top", type=int, default=20, help="展示的顶层项目/最大文件数量（默认：20）")
    parser.add_argument("--workers", type=int, default=4, help="扫描线程数（默认：4）")
    parser.add_argument("--temp-age-days", type=int, default=7, help="临时文件最小保留天数（默认：7）")
    parser.add_argument("--cache-age-days", type=int, default=30, help="缓存/诊断文件最小保留天数（默认：30）")
    return parser


def main(argv: Optional[Sequence[str]] = None) -> int:
    args = build_parser().parse_args(argv)
    if args.top < 1 or args.workers < 1 or args.temp_age_days < 0 or args.cache_age_days < 0:
        print("错误：--top/--workers 必须大于 0，保留天数不能为负数。", file=sys.stderr)
        return 2

    root = Path(args.drive)
    if not root.exists():
        print(f"错误：目标不存在：{root}", file=sys.stderr)
        return 2
    if not root.is_dir():
        print(f"错误：目标不是磁盘或目录：{root}", file=sys.stderr)
        return 2

    if args.skip_analyze and (args.full or args.scan_only):
        print("错误：--skip-analyze 不能与 --full 或 --scan-only 同时使用。", file=sys.stderr)
        return 2

    root = Path(os.path.abspath(root))
    if args.skip_analyze:
        prompt_stream = sys.stderr if args.json else sys.stdout
        confirmed = _prompt_for_direct_cleanup(
            args.temp_age_days,
            args.cache_age_days,
            stream=prompt_stream,
        )
        cleanup_result = CleanupResult(failed_paths=[])
        disk_before = shutil.disk_usage(root)
        if confirmed:
            specs = build_candidate_specs(root, args.temp_age_days, args.cache_age_days)
            try:
                with ScanProgress(enabled=not args.no_progress) as progress:
                    cleanup_result = clean_recommended_files(
                        root,
                        specs,
                        workers=args.workers,
                        now=time.time(),
                        progress=progress,
                    )
            except KeyboardInterrupt:
                print("\n清理已中断；已经删除的文件无法恢复。", file=sys.stderr)
                return 130
            if not args.json:
                _print_cleanup_result(cleanup_result)
        elif not args.json:
            print("\n已取消，未删除任何文件。")

        if args.json:
            disk_after = shutil.disk_usage(root)
            direct_report = {
                "schema_version": 1,
                "generated_at": time.strftime("%Y-%m-%dT%H:%M:%S%z"),
                "root": str(root),
                "analysis_skipped": True,
                "policy": {
                    "temp_age_days": args.temp_age_days,
                    "cache_age_days": args.cache_age_days,
                    "cleanup_requires_confirmation": "CLEAN",
                },
                "disk_before": {
                    "total_bytes": disk_before.total,
                    "used_bytes": disk_before.used,
                    "free_bytes": disk_before.free,
                },
                "disk_after": {
                    "total_bytes": disk_after.total,
                    "used_bytes": disk_after.used,
                    "free_bytes": disk_after.free,
                },
                "cleanup": _cleanup_result_dict(cleanup_result, confirmed),
            }
            json.dump(direct_report, sys.stdout, ensure_ascii=False, indent=2)
            print()
        return 0

    scan_time = time.time()
    try:
        with ScanProgress(enabled=not args.no_progress) as progress:
            report = analyze_drive(
                root,
                temp_age_days=args.temp_age_days,
                cache_age_days=args.cache_age_days,
                workers=args.workers,
                full_scan=args.full,
                largest_limit=args.top,
                now=scan_time,
                progress=progress,
            )
    except KeyboardInterrupt:
        print("\n扫描已取消。", file=sys.stderr)
        return 130
    except OSError as exc:
        print(f"错误：无法分析 {root}：{exc}", file=sys.stderr)
        return 1

    if not args.json:
        _print_report(report, args.top)

    confirmed = False
    cleanup_result = CleanupResult(failed_paths=[])
    if not args.scan_only:
        prompt_stream = sys.stderr if args.json else sys.stdout
        confirmed = _prompt_for_cleanup(report, stream=prompt_stream)
        if confirmed:
            specs = build_candidate_specs(root, args.temp_age_days, args.cache_age_days)
            try:
                with ScanProgress(enabled=not args.no_progress) as progress:
                    cleanup_result = clean_recommended_files(
                        Path(os.path.abspath(root)),
                        specs,
                        workers=args.workers,
                        now=scan_time,
                        progress=progress,
                    )
            except KeyboardInterrupt:
                print("\n清理已中断；已经删除的文件无法恢复。", file=sys.stderr)
                return 130
            if not args.json:
                _print_cleanup_result(cleanup_result)
        elif report["reclaimable_summary"][LEVEL_RECOMMENDED] > 0 and not args.json:
            print("\n已取消，未删除任何文件。")

    report["cleanup"] = _cleanup_result_dict(cleanup_result, confirmed)
    report["cleanup"]["prompted"] = not args.scan_only
    if args.json:
        json.dump(report, sys.stdout, ensure_ascii=False, indent=2)
        print()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
