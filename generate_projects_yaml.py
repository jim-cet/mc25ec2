#!/usr/bin/env python3
"""
generate_projects_yaml.py

Reads a Google-Form "Responses" .xlsx export (columns: Timestamp, Email,
Title, Batch members, Description, Google-Drive video link), downloads each
video from Google Drive, renames it to batch<N>_project.<ext> (matching the
style used in the sample projects.yaml), and writes out a new YAML file
with entries shaped like:

    - name: <title>
      batch: <batch members>
      description: <description>
      video: videos/batch<N>_project.<ext>

Install dependencies first:
    pip install gdown openpyxl pyyaml

Usage:
    python generate_projects_yaml.py responses.xlsx

Useful options:
    --outdir videos              # where downloaded videos are saved (default: videos)
    --yaml-out projects.yaml     # output YAML path (default: projects.yaml)
    --start-batch 1              # first batch number to assign (default: 1)
    --dry-run                    # show what would happen, don't download/write anything

NOTE ON BATCH NUMBERING
------------------------
The source spreadsheet has no "batch number" column, so this script assigns
batch numbers sequentially in row order (--start-batch, --start-batch+1, ...).
If you need specific batch numbers (e.g. matching a roster), edit the
BATCH_NUMBER_OVERRIDES dict below, keyed by row index (0 = first data row),
or pass --mapping a CSV with two columns "row,batch_number".

NOTE ON PRIVACY
------------------------
This version uses gdown against the plain share link, which only works for
files shared as "Anyone with the link". If your video files are private,
this will fail with a permission error -- share the files first, or ask for
the OAuth/Drive-API version of this script instead.
"""

import argparse
import csv
import re
import sys
from pathlib import Path

import openpyxl
import yaml

try:
    import gdown
except ImportError:
    gdown = None

# Optional manual overrides: {row_index (0-based, first data row = 0): batch_number}
BATCH_NUMBER_OVERRIDES = {}

DRIVE_ID_PATTERNS = [
    re.compile(r"[?&]id=([\w-]+)"),
    re.compile(r"/d/([\w-]+)"),
    re.compile(r"/file/d/([\w-]+)"),
]


def extract_drive_id(url: str):
    if not url:
        return None
    for pat in DRIVE_ID_PATTERNS:
        m = pat.search(url)
        if m:
            return m.group(1)
    return None


def load_mapping(path: str) -> dict:
    mapping = {}
    with open(path, newline="") as f:
        reader = csv.reader(f)
        for row in reader:
            if not row or not row[0].strip().isdigit():
                continue
            mapping[int(row[0])] = int(row[1])
    return mapping


def read_rows(xlsx_path: str, sheet):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb[sheet] if sheet else wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))  # skip header
    return rows


def download_video(file_id: str, dest_stub: Path, dry_run: bool):
    """Download the Drive file to dest_stub.<ext>, detecting the real
    extension from what gdown saves. Returns the final path, or None on
    failure."""
    if dry_run:
        print(f"  [dry-run] would download drive id={file_id} -> {dest_stub}.*")
        return dest_stub.with_suffix(".mp4")

    if gdown is None:
        print("  ERROR: gdown is not installed. Run: pip install gdown", file=sys.stderr)
        return None

    tmp_target = str(dest_stub) + ".download"
    try:
        result = gdown.download(id=file_id, output=tmp_target, quiet=False)
    except Exception as e:
        print(f"  ERROR downloading {file_id}: {e}", file=sys.stderr)
        return None

    if not result:
        print(f"  ERROR: gdown returned nothing for {file_id}", file=sys.stderr)
        return None

    downloaded_path = Path(result)
    ext = downloaded_path.suffix or ".mp4"
    final_path = dest_stub.with_suffix(ext)
    downloaded_path.rename(final_path)
    return final_path


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("xlsx", help="Path to the Form Responses .xlsx file")
    ap.add_argument("--sheet", default=None, help="Worksheet name (default: active sheet)")
    ap.add_argument("--outdir", default="videos", help="Directory to save downloaded videos (default: videos)")
    ap.add_argument("--yaml-out", default="projects.yaml", help="Output YAML path (default: projects.yaml)")
    ap.add_argument("--start-batch", type=int, default=1, help="First batch number to assign (default: 1)")
    ap.add_argument("--mapping", default=None, help="Optional CSV file: row_index,batch_number (row 0 = first data row)")
    ap.add_argument("--dry-run", action="store_true", help="Print planned actions without downloading or writing files")
    args = ap.parse_args()

    if args.mapping:
        BATCH_NUMBER_OVERRIDES.update(load_mapping(args.mapping))

    outdir = Path(args.outdir)
    if not args.dry_run:
        outdir.mkdir(parents=True, exist_ok=True)

    rows = read_rows(args.xlsx, args.sheet)

    entries = []
    next_batch_num = args.start_batch

    for idx, row in enumerate(rows):
        if row is None or all(c is None for c in row):
            continue

        # Expected columns: Timestamp, Email, Title, Batch members, Description, Video link
        _, _, title, batch_members, description, video_url = (list(row) + [None] * 6)[:6]

        title = (title or "").strip()
        batch_members = (batch_members or "").strip()
        description = (description or "").strip()
        video_url = (video_url or "").strip()

        if not title or not video_url:
            print(f"Row {idx}: skipping (missing title or video link)")
            continue

        batch_num = BATCH_NUMBER_OVERRIDES.get(idx, next_batch_num)
        if idx not in BATCH_NUMBER_OVERRIDES:
            next_batch_num += 1

        print(f"Row {idx}: '{title}' -> batch{batch_num}")

        file_id = extract_drive_id(video_url)
        if not file_id:
            print(f"  WARNING: could not parse Drive file id from '{video_url}', skipping download", file=sys.stderr)
            video_rel_path = f"videos/batch{batch_num}_project.mp4  # MANUAL: could not resolve link"
        else:
            dest_stub = outdir / f"batch{batch_num}_project"
            final_path = download_video(file_id, dest_stub, args.dry_run)
            if final_path is None:
                video_rel_path = f"videos/batch{batch_num}_project.mp4  # MANUAL: download failed"
            else:
                video_rel_path = f"videos/{final_path.name}"

        entries.append({
            "name": title,
            "batch": batch_members,
            "description": description,
            "video": video_rel_path,
        })

    if args.dry_run:
        print(f"\n[dry-run] Would write {len(entries)} entries to {args.yaml_out}")
        return

    with open(args.yaml_out, "w") as f:
        yaml.safe_dump(entries, f, sort_keys=False, allow_unicode=True, width=1000)

    print(f"\nWrote {len(entries)} entries to {args.yaml_out}")
    print(f"Videos saved under {outdir}/")


if __name__ == "__main__":
    main()
