"""Download every PDF listed in data/catalog.json to its normalized file name.

Usage: python scripts/download_pdfs.py [--force]
Writes data/download_report.json; re-run scripts/build_table.py afterwards
so the table reflects download status. Existing files are kept unless --force.
"""

from __future__ import annotations

import hashlib
import http.client
import json
import logging
import sys
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any
from urllib.parse import quote, urlsplit, urlunsplit

ROOT = Path(__file__).resolve().parent.parent
CATALOG_FILE = ROOT / "data" / "catalog.json"
REPORT_FILE = ROOT / "data" / "download_report.json"
MAX_BYTES = 50 * 1024 * 1024
TIMEOUT_S = 60
WORKERS = 8
RETRY_DELAY_S = 3
PDF_MAGIC_WINDOW = 1024
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
)
# Some hosts (e.g. agsc.org) drop connections for a full browser UA but accept a short one.
FALLBACK_USER_AGENT = "Mozilla/5.0"
RETRY_STATUSES = {"network_error", "http_403"}

log = logging.getLogger("download_pdfs")


def encode_url(url: str) -> str:
    parts = urlsplit(url)
    return urlunsplit(parts._replace(path=quote(parts.path, safe="/%:@+-._~")))


def fetch(url: str, user_agent: str = USER_AGENT) -> tuple[str, bytes, dict[str, Any]]:
    """Return (status, payload, meta); status is 'ok' only for a real PDF body."""
    if not url.lower().startswith(("https://", "http://")):
        return "invalid_url", b"", {}
    request = urllib.request.Request(
        encode_url(url),
        headers={"User-Agent": user_agent, "Accept": "application/pdf,*/*;q=0.8"},
    )
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT_S) as response:
            payload = response.read(MAX_BYTES + 1)
            meta = {"final_url": response.geturl(),
                    "content_type": response.headers.get("Content-Type", "")}
    except urllib.error.HTTPError as exc:
        return f"http_{exc.code}", b"", {}
    except (urllib.error.URLError, http.client.HTTPException, TimeoutError,
            ConnectionError) as exc:
        return "network_error", b"", {"error": str(exc)}
    if len(payload) > MAX_BYTES:
        return "too_large", b"", meta
    if b"%PDF" not in payload[:PDF_MAGIC_WINDOW]:
        return "not_pdf", b"", meta
    return "ok", payload, meta


def load_previous() -> dict[str, dict[str, Any]]:
    if not REPORT_FILE.exists():
        return {}
    return {row["file"]: row for row in json.loads(REPORT_FILE.read_text(encoding="utf-8"))}


def download(rel_path: str, url: str, force: bool, previous: dict[str, dict[str, Any]]) -> dict[str, Any]:
    """Reuse a local file only if it was downloaded from this same URL and is unchanged."""
    target = ROOT / rel_path
    prev = previous.get(rel_path, {})
    if target.exists() and not force and prev.get("url") == url and prev.get("status") == "ok":
        data = target.read_bytes()
        if hashlib.sha256(data).hexdigest() == prev.get("sha256"):
            return prev
    status, payload, meta = fetch(url)
    if status in RETRY_STATUSES:
        time.sleep(RETRY_DELAY_S)
        status, payload, meta = fetch(url, FALLBACK_USER_AGENT)
    row: dict[str, Any] = {"file": rel_path, "url": url, "status": status, **meta}
    if status == "ok":
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(payload)
        row.update(bytes=len(payload), sha256=hashlib.sha256(payload).hexdigest())
    else:
        log.warning("%-14s %s", status, url)
    return row


def reuse_renamed(jobs: dict[str, str], previous: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    """Move files whose catalog name changed (same URL, same content) instead of re-downloading."""
    by_url = {row["url"]: row for row in previous.values() if row.get("status") == "ok"}
    updated = dict(previous)
    for rel_path, url in jobs.items():
        old = by_url.get(url)
        if not old or old["file"] == rel_path:
            continue
        source, target = ROOT / old["file"], ROOT / rel_path
        if source.exists() and not target.exists() and \
                hashlib.sha256(source.read_bytes()).hexdigest() == old.get("sha256"):
            target.parent.mkdir(parents=True, exist_ok=True)
            source.rename(target)
            updated[rel_path] = {**old, "file": rel_path}
    return updated


def remove_stale(expected: set[Path]) -> None:
    """Delete PDFs left over from earlier catalog versions (pdf/ is a generated cache)."""
    for folder in {path.parent for path in expected}:
        for path in folder.glob("*.pdf"):
            if path not in expected:
                log.info("removing stale %s", path.relative_to(ROOT))
                path.unlink()


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    force = "--force" in sys.argv[1:]
    catalog = json.loads(CATALOG_FILE.read_text(encoding="utf-8"))
    jobs: dict[str, str] = {}
    for item in catalog:
        for doc in item["material_docs"] + item["equipment_docs"]:
            if doc["file"]:
                jobs.setdefault(doc["file"], doc["url"])

    previous = load_previous()
    previous = reuse_renamed(jobs, previous)
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        report = list(pool.map(lambda job: download(*job, force, previous), jobs.items()))
    remove_stale({ROOT / rel for rel in jobs})
    REPORT_FILE.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    ok = sum(r["status"] == "ok" for r in report)
    print(f"{ok}/{len(report)} downloaded")


if __name__ == "__main__":
    main()
