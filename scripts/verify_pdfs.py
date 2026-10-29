"""Check that each downloaded PDF really is the document the catalog claims.

Usage: python scripts/verify_pdfs.py <excerpt_dir>
Reads data/catalog.json + data/download_report.json, extracts text with poppler
(pdfinfo/pdftotext), scores catalog keywords against the text and writes
data/verification.json. Text excerpts go to <excerpt_dir> for manual review.
"""

from __future__ import annotations

import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
CATALOG_FILE = ROOT / "data" / "catalog.json"
REPORT_FILE = ROOT / "data" / "download_report.json"
OUT_FILE = ROOT / "data" / "verification.json"
TEXT_PAGES = 3
EXCERPT_CHARS = 1500
TOOL_TIMEOUT_S = 60
STOPWORDS = {
    "the", "and", "for", "with", "pdf", "sds", "tds", "pds", "msds", "data", "sheet",
    "safety", "technical", "product", "brochure", "manual", "datasheet", "information",
    "series", "system", "systems", "guide", "instructions", "operation", "en", "eng",
    "english", "official", "copy", "mirror", "version", "rev", "edition", "of", "type",
    "page", "html", "new", "kit", "unit", "units", "automotive", "oem", "v", "us", "uk",
}


def run(cmd: list[str]) -> str:
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=TOOL_TIMEOUT_S,
                                check=False, errors="replace")
    except subprocess.TimeoutExpired:
        return ""
    return result.stdout


def pdf_info(path: Path) -> dict[str, str]:
    info: dict[str, str] = {}
    for line in run(["pdfinfo", str(path)]).splitlines():
        key, _, value = line.partition(":")
        if key in {"Title", "Pages", "Producer", "CreationDate"}:
            info[key.lower()] = value.strip()
    return info


def normalize(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", text.lower())


def keywords(doc: dict[str, Any]) -> list[str]:
    source = " ".join(str(doc.get(k, "")) for k in ("title", "model", "vendor"))
    tokens = [t for t in normalize(source).split() if t not in STOPWORDS and len(t) > 1]
    return list(dict.fromkeys(tokens))


def score(tokens: list[str], text: str) -> tuple[float, list[str]]:
    haystack = f" {normalize(text)} "
    missing = [t for t in tokens if f" {t} " not in haystack]
    hit_ratio = 1 - len(missing) / len(tokens) if tokens else 0.0
    return round(hit_ratio, 2), missing


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    excerpt_dir = Path(sys.argv[1])
    excerpt_dir.mkdir(parents=True, exist_ok=True)
    status = {r["file"]: r for r in json.loads(REPORT_FILE.read_text(encoding="utf-8"))}
    catalog = json.loads(CATALOG_FILE.read_text(encoding="utf-8"))

    docs: dict[str, dict[str, Any]] = {}
    for item in catalog:
        for kind in ("material_docs", "equipment_docs"):
            for doc in item[kind]:
                if doc["file"]:
                    entry = docs.setdefault(doc["file"], {"doc": doc, "kind": kind, "ids": []})
                    entry["ids"].append(item["id"])

    results = []
    for rel_path, entry in docs.items():
        row: dict[str, Any] = {"file": rel_path, "ids": entry["ids"], "kind": entry["kind"],
                               "url": entry["doc"]["url"], "title": entry["doc"].get("title", ""),
                               "download": status.get(rel_path, {}).get("status", "missing")}
        path = ROOT / rel_path
        if row["download"] == "ok" and path.exists():
            info = pdf_info(path)
            text = run(["pdftotext", "-l", str(TEXT_PAGES), "-layout", str(path), "-"])
            tokens = keywords(entry["doc"])
            ratio, missing = score(tokens, f"{info.get('title', '')} {text}")
            row.update(pages=info.get("pages"), pdf_title=info.get("title", ""),
                       text_chars=len(text.strip()), keyword_ratio=ratio, missing=missing)
            excerpt = re.sub(r"\s+", " ", text).strip()[:EXCERPT_CHARS]
            (excerpt_dir / f"{path.stem}.txt").write_text(excerpt, encoding="utf-8")
        results.append(row)
    OUT_FILE.write_text(json.dumps(results, ensure_ascii=False, indent=1), encoding="utf-8")
    checked = [r for r in results if "keyword_ratio" in r]
    weak = [r for r in checked if r["keyword_ratio"] < 0.6 or r["text_chars"] < 200]
    print(f"{len(checked)} checked, {len(weak)} need manual review -> {OUT_FILE.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
