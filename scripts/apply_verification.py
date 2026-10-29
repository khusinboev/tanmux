"""Apply PDF content-verification verdicts to data/research/group_*.json.

Usage: python scripts/apply_verification.py <verify.json> [<verify.json> ...]
Each verify file is a JSON array of rows with: url, action (keep|relabel|remove),
suggested_title, actual_document, reason, and for material docs label_should_be.
Kept/relabelled docs get verified=true and the real title; removed docs are dropped.
A log of every change is appended to data/verification_log.json.
"""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
RESEARCH_DIR = ROOT / "data" / "research"
LOG_FILE = ROOT / "data" / "verification_log.json"
DOC_FIELDS = ("material_docs", "equipment_docs")
ACTIONS = {"keep", "relabel", "remove"}

log = logging.getLogger("apply_verification")


def load_rows(paths: list[Path]) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    for path in paths:
        for row in json.loads(path.read_text(encoding="utf-8")):
            if row.get("action") not in ACTIONS:
                raise SystemExit(f"{path.name}: bad action {row.get('action')!r} for {row.get('url')}")
            rows[row["url"]] = row
    return rows


def apply_row(doc: dict[str, Any], row: dict[str, Any]) -> None:
    doc["verified"] = True
    if row.get("suggested_title"):
        doc["title"] = row["suggested_title"]
    label = row.get("label_should_be")
    if label in {"exact", "equivalent"}:
        doc["exact_or_equivalent"] = label
    doc["verification"] = {
        key: row[key]
        for key in ("verdict", "actual_document", "doc_type", "language", "version_date",
                    "doc_number_rev", "class_ok", "passport_ok", "process_ok", "reason")
        if key in row
    }


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    rows = load_rows([Path(p) for p in sys.argv[1:]])
    changes: list[dict[str, Any]] = []
    seen: set[str] = set()
    for path in sorted(RESEARCH_DIR.glob("group_*.json")):
        entries = json.loads(path.read_text(encoding="utf-8"))
        for entry in entries:
            for field in DOC_FIELDS:
                kept = []
                for doc in entry.get(field, []):
                    row = rows.get(doc.get("url", ""))
                    if row is None:
                        kept.append(doc)
                        continue
                    seen.add(doc["url"])
                    changes.append({"id": entry["id"], "field": field, "url": doc["url"],
                                    "action": row["action"], "reason": row.get("reason", "")})
                    if row["action"] != "remove":
                        apply_row(doc, row)
                        kept.append(doc)
                entry[field] = kept
        path.write_text(json.dumps(entries, ensure_ascii=False, indent=1), encoding="utf-8")

    unmatched = sorted(set(rows) - seen)
    if unmatched:
        log.warning("%d verification rows matched no catalog doc: %s", len(unmatched), unmatched[:5])
    history = json.loads(LOG_FILE.read_text(encoding="utf-8")) if LOG_FILE.exists() else []
    LOG_FILE.write_text(json.dumps(history + changes, ensure_ascii=False, indent=1), encoding="utf-8")
    counts = {a: sum(c["action"] == a for c in changes) for a in sorted(ACTIONS)}
    print(f"applied {len(changes)} doc changes {counts}")


if __name__ == "__main__":
    main()
