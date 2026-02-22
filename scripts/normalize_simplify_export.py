#!/usr/bin/env python3
"""Normalize Simplify export CSV into a stable schema for job funnel ingestion."""

from __future__ import annotations

import argparse
import csv
from datetime import datetime
from pathlib import Path
from typing import Dict, List

TARGET_COLUMNS = [
    "company",
    "role",
    "location",
    "applied_date",
    "status",
    "job_url",
    "source",
]

FIELD_ALIASES = {
    "company": ["company", "company name"],
    "role": ["role", "job title", "title", "position"],
    "location": ["location", "job location"],
    "applied_date": ["date applied", "applied date", "application date"],
    "status": ["status", "application status"],
    "job_url": ["job url", "posting url", "url", "link"],
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--input", required=True, help="Path to Simplify export CSV")
    parser.add_argument("--output", required=True, help="Path for normalized CSV")
    parser.add_argument(
        "--source",
        default="Simplify",
        help="Source label to embed in output rows (default: Simplify)",
    )
    return parser.parse_args()


def normalize_header(row: Dict[str, str]) -> Dict[str, str]:
    out = {k.strip().lower(): (v or "").strip() for k, v in row.items() if k}
    return out


def pick(row: Dict[str, str], aliases: List[str]) -> str:
    for alias in aliases:
        if alias in row and row[alias]:
            return row[alias]
    return ""


def normalize_date(value: str) -> str:
    value = value.strip()
    if not value:
        return ""

    date_formats = [
        "%Y-%m-%d",
        "%m/%d/%Y",
        "%m/%d/%y",
        "%b %d, %Y",
        "%B %d, %Y",
    ]
    for fmt in date_formats:
        try:
            dt = datetime.strptime(value, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            pass

    try:
        dt = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return dt.strftime("%Y-%m-%d")
    except ValueError:
        return value


def normalize_row(raw_row: Dict[str, str], source: str) -> Dict[str, str]:
    row = normalize_header(raw_row)
    normalized = {
        "company": pick(row, FIELD_ALIASES["company"]),
        "role": pick(row, FIELD_ALIASES["role"]),
        "location": pick(row, FIELD_ALIASES["location"]),
        "applied_date": normalize_date(pick(row, FIELD_ALIASES["applied_date"])),
        "status": pick(row, FIELD_ALIASES["status"]) or "Applied",
        "job_url": pick(row, FIELD_ALIASES["job_url"]),
        "source": source,
    }
    return normalized


def main() -> None:
    args = parse_args()

    input_path = Path(args.input)
    output_path = Path(args.output)

    if not input_path.exists():
        raise FileNotFoundError(f"Input CSV not found: {input_path}")

    with input_path.open("r", newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        rows = [normalize_row(row, args.source) for row in reader]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=TARGET_COLUMNS)
        writer.writeheader()
        writer.writerows(rows)

    print(f"Normalized {len(rows)} rows -> {output_path}")


if __name__ == "__main__":
    main()
