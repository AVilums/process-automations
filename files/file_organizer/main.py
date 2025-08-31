import argparse
import sys
from pathlib import Path
import shutil
import os
from datetime import datetime

# Mapping of file extensions to folder names
EXTENSION_MAP = {
    'images': {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.svg', '.webp', '.heic'},
    'documents': {'.pdf', '.doc', '.docx', '.txt', '.rtf', '.odt', '.xls', '.xlsx', '.ppt', '.pptx', '.csv', '.md'},
    'videos': {'.mp4', '.mov', '.avi', '.mkv', '.wmv', '.flv', '.webm', '.m4v'},
    'audio': {'.mp3', '.wav', '.aac', '.flac', '.ogg', '.m4a', '.wma'},
    'archives': {'.zip', '.rar', '.7z', '.tar', '.gz', '.bz2', '.xz'},
    'code': {'.py', '.js', '.ts', '.java', '.c', '.cpp', '.cs', '.rb', '.go', '.php', '.html', '.css', '.json', '.xml', '.yml', '.yaml', '.ini', '.sh', '.bat', '.ps1'},
    'executables': {'.exe', '.msi', '.apk', '.app', '.deb', '.rpm'}
}

CATEGORY_FOLDERS = {
    'images': 'Images',
    'documents': 'Documents',
    'videos': 'Videos',
    'audio': 'Audio',
    'archives': 'Archives',
    'code': 'Code',
    'executables': 'Executables',
    'others': 'Others',
}

def categorize(file_path: Path) -> str:
    ext = file_path.suffix.lower()
    for category, exts in EXTENSION_MAP.items():
        if ext in exts:
            return category
    return 'others'


def ensure_directory(path: Path):
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)


def move_file(src: Path, dest_dir: Path, dry_run: bool) -> None:
    ensure_directory(dest_dir)
    target = dest_dir / src.name

    # Avoid overwriting: if exists, append a counter
    if target.exists():
        stem, suffix = src.stem, src.suffix
        counter = 1
        while True:
            candidate = dest_dir / f"{stem} ({counter}){suffix}"
            if not candidate.exists():
                target = candidate
                break
            counter += 1

    if dry_run:
        print(f"[DRY-RUN] Would move: {src} -> {target}")
        return

    shutil.move(str(src), str(target))
    print(f"Moved: {src} -> {target}")


def organize_directory(source_dir: Path, dry_run: bool) -> int:
    if not source_dir.exists() or not source_dir.is_dir():
        print(f"Error: Source directory does not exist or is not a directory: {source_dir}", file=sys.stderr)
        return 1

    print(f"Starting organization at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\nDirectory: {source_dir}\nDry-run: {dry_run}")

    # Skip our own category folders to avoid recursing/moving them
    skip_dirs = {CATEGORY_FOLDERS[k].lower() for k in CATEGORY_FOLDERS}

    exit_code = 0
    try:
        for entry in source_dir.iterdir():
            # Skip directories (we only organize top-level files)
            if entry.is_dir():
                # But don't skip if it's a directory with extension like ".config" (treat as dir)
                # We simply skip all directories to avoid moving nested content unexpectedly
                continue

            if not entry.is_file():
                continue

            # Skip if file is already inside a category folder (not applicable since we don't recurse)
            category = categorize(entry)
            dest_folder_name = CATEGORY_FOLDERS.get(category, 'Others')
            dest_dir = source_dir / dest_folder_name

            # Ensure we don't move the file into itself or into one of the skip dirs incorrectly
            if entry.name.lower() in skip_dirs:
                continue

            try:
                move_file(entry, dest_dir, dry_run)
            except PermissionError as e:
                print(f"Permission denied moving {entry}: {e}", file=sys.stderr)
                exit_code = 1
            except OSError as e:
                print(f"OS error moving {entry}: {e}", file=sys.stderr)
                exit_code = 1
    except PermissionError as e:
        print(f"Permission denied accessing directory {source_dir}: {e}", file=sys.stderr)
        exit_code = 1
    except OSError as e:
        print(f"OS error accessing directory {source_dir}: {e}", file=sys.stderr)
        exit_code = 1

    if exit_code == 0:
        print("Organization completed successfully.")
    else:
        print("Organization completed with errors.")
    return exit_code


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="Organize files in a directory into categorized subfolders by extension.")
    parser.add_argument('--source-dir', required=True, help='Path to the directory to organize')
    parser.add_argument('--dry-run', action='store_true', help='Show what would be done without moving files')
    return parser.parse_args(argv)


def main(argv=None):
    args = parse_args(argv)
    source_dir = Path(args.source_dir).expanduser()
    try:
        code = organize_directory(source_dir, args.dry_run)
    except Exception as e:
        print(f"Unexpected error: {e}", file=sys.stderr)
        return 1
    return code


if __name__ == '__main__':
    sys.exit(main())
