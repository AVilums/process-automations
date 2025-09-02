import argparse
import sys
import os
import shutil
import subprocess
from pathlib import Path
from datetime import datetime, timedelta
import platform


def get_windows_drives():
    drives = []
    for letter in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
        path = f"{letter}:\\"
        if os.path.exists(path):
            drives.append(path)
    return drives


def collect_disk_usage():
    items = []
    if os.name == 'nt':
        paths = get_windows_drives()
    else:
        # On POSIX, check common mount points (root sufficient for simple report)
        paths = ['/']
    for p in paths:
        try:
            total, used, free = shutil.disk_usage(p)
            percent_used = (used / total * 100.0) if total > 0 else 0.0
            items.append({
                'mount': p,
                'total': total,
                'used': used,
                'free': free,
                'percent_used': percent_used,
            })
        except PermissionError:
            # Skip inaccessible mount points
            continue
        except OSError:
            continue
    return items


def count_processes():
    print('Counting processes...')
    try:
        if os.name == 'nt':
            # Use tasklist and count non-header lines
            result = subprocess.run(['cmd', '/c', 'tasklist'], capture_output=True, text=True, timeout=10)
            if result.returncode != 0:
                return None
            lines = [l for l in result.stdout.splitlines() if l.strip()]
            # tasklist has a header and a separator line; find the separator of === and count after it
            count = 0
            sep_found = False
            for line in lines:
                if not sep_found:
                    if set(line.strip()) == {'='}:
                        sep_found = True
                    continue
                else:
                    count += 1
            # Fallback: if separator not found, try excluding first 3 lines
            if not sep_found:
                count = max(0, len(lines) - 3)
            return count
        else:
            result = subprocess.run(['ps', '-e'], capture_output=True, text=True, timeout=10)
            if result.returncode != 0:
                return None
            lines = [l for l in result.stdout.splitlines() if l.strip()]
            # Exclude header
            return max(0, len(lines) - 1)
    except Exception:
        return None


def get_uptime():
    if os.name == 'nt':
        try:
            import ctypes
            GetTickCount64 = ctypes.windll.kernel32.GetTickCount64
            GetTickCount64.restype = ctypes.c_ulonglong
            ms = GetTickCount64()
            return timedelta(milliseconds=int(ms))
        except Exception:
            return None
    else:
        try:
            with open('/proc/uptime', 'r') as f:
                contents = f.read().strip().split()
                seconds = float(contents[0])
                return timedelta(seconds=int(seconds))
        except Exception:
            return None


def severity_from_disks(disks):
    # Return 2 if any > 90%, 1 if any > 80%, else 0
    sev = 0
    for d in disks:
        pu = d.get('percent_used', 0.0)
        if pu > 90.0:
            return 2
        if pu > 80.0:
            sev = max(sev, 1)
    return sev


def human_bytes(n: int) -> str:
    # Human-readable bytes
    step = 1024.0
    units = ['B', 'KB', 'MB', 'GB', 'TB', 'PB']
    size = float(n)
    for u in units:
        if size < step:
            return f"{size:,.0f} {u}" if u == 'B' else f"{size:,.2f} {u}"
        size /= step
    return f"{size:,.2f} EB"


def build_report(disks, proc_count, uptime_td):
    lines = []
    lines.append(f"System Health Report - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append(f"OS: {platform.system()} {platform.release()} ({platform.version()})")
    lines.append("")
    lines.append("Disk Usage:")
    if not disks:
        lines.append("  No disks found or accessible.")
    else:
        for d in disks:
            lines.append(
                f"  {d['mount']}: used {human_bytes(d['used'])} / {human_bytes(d['total'])} "
                f"({d['percent_used']:.1f}% used, {human_bytes(d['free'])} free)"
            )
    lines.append("")
    lines.append(f"Running processes: {proc_count if proc_count is not None else 'N/A'}")
    lines.append(f"System uptime: {str(uptime_td) if uptime_td is not None else 'N/A'}")
    return "\n".join(lines)


def write_or_print(report: str, output_file: str | None) -> int:
    if output_file:
        try:
            out_path = Path(output_file).expanduser()
            out_path.parent.mkdir(parents=True, exist_ok=True)
            out_path.write_text(report, encoding='utf-8')
            print(f"Report written to: {out_path}")
        except Exception as e:
            print(f"Failed to write report: {e}", file=sys.stderr)
            return 2
    else:
        print(report)
    return 0


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description='Perform basic system health checks and generate a report.')
    parser.add_argument('--output-file', help='Path to write the report (optional). If omitted, prints to console.')
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    disks = collect_disk_usage()
    proc_count = count_processes()
    uptime_td = get_uptime()

    report = build_report(disks, proc_count, uptime_td)

    # Determine exit code from disk usage thresholds only (per requirements)
    sev = severity_from_disks(disks)

    write_rc = write_or_print(report, args.output_file)
    if write_rc != 0:
        return write_rc

    return sev


if __name__ == '__main__':
    sys.exit(main())
