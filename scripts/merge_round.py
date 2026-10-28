"""Merge an additional research round into data/research/group_*.json.

Usage: python scripts/merge_round.py <round_file.json> [<round_file.json> ...]
Each round file is a JSON array of objects with keys: id, manufacturer_update,
add_material_docs, add_equipment, add_equipment_docs, remove_urls, notes.
"""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parent.parent
RESEARCH_DIR = ROOT / "data" / "research"

log = logging.getLogger("merge_round")


def load_groups() -> tuple[dict[Path, list[dict[str, Any]]], dict[int, dict[str, Any]]]:
    groups = {p: json.loads(p.read_text(encoding="utf-8")) for p in sorted(RESEARCH_DIR.glob("group_*.json"))}
    by_id = {int(e["id"]): e for entries in groups.values() for e in entries}
    return groups, by_id


def append_unique(target: list[dict[str, Any]], additions: list[dict[str, Any]], key: str) -> int:
    seen = {d.get(key) for d in target}
    added = 0
    for doc in additions:
        if not doc.get(key) or doc[key] in seen:
            continue
        target.append(doc)
        seen.add(doc[key])
        added += 1
    return added


def equipment_key(e: dict[str, Any]) -> str:
    return f"{e.get('vendor', '')}|{e.get('model', '')}".lower()


def apply_entry(entry: dict[str, Any], update: dict[str, Any], round_name: str) -> dict[str, int]:
    removed = {r["url"] for r in update.get("remove_urls") or [] if r.get("url")}
    for field in ("material_docs", "equipment_docs"):
        entry[field] = [d for d in entry.get(field, []) if d.get("url") not in removed]

    stats = {
        "material_docs": append_unique(entry.setdefault("material_docs", []),
                                       update.get("add_material_docs") or [], "url"),
        "equipment_docs": append_unique(entry.setdefault("equipment_docs", []),
                                        update.get("add_equipment_docs") or [], "url"),
        "removed": len(removed),
    }
    existing = {equipment_key(e) for e in entry.setdefault("equipment", [])}
    for eq in update.get("add_equipment") or []:
        if equipment_key(eq) not in existing:
            entry["equipment"].append(eq)
            existing.add(equipment_key(eq))

    manufacturer = update.get("manufacturer_update")
    if manufacturer:
        for field in ("manufacturer", "product_identified", "manufacturer_confidence"):
            if manufacturer.get(field):
                entry[field] = manufacturer[field]
        if manufacturer.get("evidence"):
            entry["notes"] = f"{entry.get('notes', '')} | {round_name} evidence: {manufacturer['evidence']}".strip(" |")
    if update.get("notes"):
        entry["notes"] = f"{entry.get('notes', '')} | {round_name}: {update['notes']}".strip(" |")
    return stats


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    if len(sys.argv) < 2:
        raise SystemExit(__doc__)
    groups, by_id = load_groups()
    for path in map(Path, sys.argv[1:]):
        round_name = path.stem
        for update in json.loads(path.read_text(encoding="utf-8")):
            item_id = int(update["id"])
            if item_id not in by_id:
                log.error("%s: unknown material id %s, skipped", round_name, item_id)
                continue
            stats = apply_entry(by_id[item_id], update, round_name)
            log.info("%s id=%s %s", round_name, item_id, stats)
    for path, entries in groups.items():
        path.write_text(json.dumps(entries, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
