"""Download every PDF listed in data/catalog.json to its normalized file name.

Usage: python scripts/download_pdfs.py
Writes data/download_report.json; re-run scripts/build_table.py afterwards
so the table reflects download status.
"""

from __future__ import annotations

import json
import logging
import urllib.error
import urllib.request
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CATALOG_FILE = ROOT / "data" / "catalog.json"
REPORT_FILE = ROOT / "data" / "download_report.json"
MAX_BYTES = 50 * 1024 * 1024
TIMEOUT_S = 60
USER_AGENT = "Mozilla/5.0 (X11; Linux x86_64) tanmux-pdf-fetcher/1.0"

log = logging.getLogger("download_pdfs")


def fetch_pdf(url: str, target: Path) -> str:
    if not url.lower().startswith("https://") and not url.lower().startswith("http://"):
        return "invalid_url"
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(request, timeout=TIMEOUT_S) as response:
            payload = response.read(MAX_BYTES + 1)
    except urllib.error.HTTPError as exc:
        log.warning("HTTP %s for %s", exc.code, url)
        return f"http_{exc.code}"
    except (urllib.error.URLError, TimeoutError, ConnectionError) as exc:
        log.warning("Network error for %s: %s", url, exc)
        return "network_error"
    if len(payload) > MAX_BYTES:
        return "too_large"
    if not payload.startswith(b"%PDF"):
        log.warning("Not a PDF: %s", url)
        return "not_pdf"
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_bytes(payload)
    return "ok"


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    catalog = json.loads(CATALOG_FILE.read_text(encoding="utf-8"))
    jobs: dict[str, str] = {}
    for item in catalog:
        for doc in item["material_docs"] + item["equipment_docs"]:
            jobs.setdefault(doc["file"], doc["url"])

    report = []
    for rel_path, url in jobs.items():
        target = ROOT / rel_path
        status = "ok" if target.exists() else fetch_pdf(url, target)
        log.info("%-12s %s", status, rel_path)
        report.append({"file": rel_path, "url": url, "status": status})
    REPORT_FILE.write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    ok = sum(r["status"] == "ok" for r in report)
    print(f"{ok}/{len(report)} downloaded")


if __name__ == "__main__":
    main()
