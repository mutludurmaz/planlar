#!/usr/bin/env python3
"""planlar deposundaki .xlsx dosyalarindan catalog.json uretir."""

from __future__ import annotations

import json
import re
import unicodedata
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parents[1]
REPO = "mutludurmaz/planlar"
BRANCH = "main"
RAW_BASE = f"https://raw.githubusercontent.com/{REPO}/{BRANCH}"
OUT = ROOT / "catalog.json"
GRADE_RE = re.compile(r"(?<!\d)(9|10|11|12)(?!\d)")


def fold(text: str) -> str:
    n = unicodedata.normalize("NFKC", text).casefold()
    return (
        n.replace("ı", "i")
        .replace("i̇", "i")
        .replace("â", "a")
        .replace("ê", "e")
        .replace("î", "i")
        .replace("û", "u")
    )


def school_type(file_stem: str) -> tuple[str, str]:
    n = fold(file_stem)
    if "fen lise" in n:
        return "fen", "Fen Lisesi"
    if "anadolu lise" in n:
        return "anadolu", "Anadolu Lisesi"
    return "diger", file_stem.strip()


def posix_path(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def raw_url(rel: str) -> str:
    encoded = "/".join(quote(part) for part in rel.split("/"))
    return f"{RAW_BASE}/{encoded}"


def grades_from_sheets(path: Path) -> list[int]:
    try:
        from openpyxl import load_workbook
    except ImportError:
        return [9, 10, 11, 12]
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        names = list(wb.sheetnames)
        wb.close()
    except Exception:
        return [9, 10, 11, 12]
    found: list[int] = []
    for name in names:
        match = GRADE_RE.search(name)
        if not match:
            continue
        grade = int(match.group(1))
        if grade not in found:
            found.append(grade)
    found.sort()
    return found or [9, 10, 11, 12]


def build() -> dict:
    by_subject: dict[str, dict] = {}
    files = sorted(
        (p for p in ROOT.rglob("*.xlsx") if ".git" not in p.parts),
        key=lambda p: (fold(p.parent.name), fold(p.name)),
    )
    for path in files:
        if path.parent == ROOT:
            continue
        subject = path.parent.name
        school_id, school_label = school_type(path.stem)
        rel = posix_path(path)
        school = {
            "id": school_id,
            "title": school_label,
            "fileName": path.name,
            "path": rel,
            "downloadUrl": raw_url(rel),
            "grades": grades_from_sheets(path),
        }
        bucket = by_subject.setdefault(
            subject,
            {"title": subject, "path": path.parent.relative_to(ROOT).as_posix(), "schools": []},
        )
        bucket["schools"].append(school)

    subjects = list(by_subject.values())
    subjects.sort(key=lambda item: fold(item["title"]))
    for subject in subjects:
        subject["schools"].sort(
            key=lambda item: (
                0 if item["id"] == "anadolu" else 1 if item["id"] == "fen" else 2,
                fold(item["title"]),
            )
        )
    return {
        "version": 1,
        "generatedAt": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "source": f"https://github.com/{REPO}",
        "subjects": subjects,
    }


def main() -> None:
    catalog = build()
    OUT.write_text(
        json.dumps(catalog, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    count = sum(len(s["schools"]) for s in catalog["subjects"])
    print(f"catalog.json: {len(catalog['subjects'])} ders, {count} dosya")


if __name__ == "__main__":
    main()
