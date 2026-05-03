#!/usr/bin/env python3
"""
download_logos.py — Download company logos to assets/logos/ from Hunter.io.

Hunter.io's logo API (https://logos.hunter.io/{domain}) returns a 128px PNG
for any domain. No API key, no rate limit, free.

Usage:
    # Download a single logo
    python scripts/download_logos.py microsoft.com

    # Download many at once
    python scripts/download_logos.py microsoft.com aws.amazon.com cloud.google.com

    # Or read from a file (one domain per line)
    python scripts/download_logos.py --file companies.txt

The logo is saved as assets/logos/{domain}.png (relative to the repo root).
The build_deck.py script auto-downloads any logo it does not find in the cache,
so running this manually is only useful for pre-warming the cache or for
visibility into what got downloaded.
"""

from __future__ import annotations

import sys
import urllib.request
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
LOGO_DIR = REPO_ROOT / "assets" / "logos"
HUNTER_URL = "https://logos.hunter.io/{domain}"
MIN_BYTES = 200  # smaller than this is probably an error response, not a logo


def safe_name(domain: str) -> str:
    return domain.strip().lower().replace("/", "_")


def fetch_logo(domain: str, force: bool = False) -> Path | None:
    LOGO_DIR.mkdir(parents=True, exist_ok=True)
    out = LOGO_DIR / f"{safe_name(domain)}.png"
    if out.exists() and not force:
        return out
    url = HUNTER_URL.format(domain=domain)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "mbb-decks/1.0"})
        with urllib.request.urlopen(req, timeout=10) as resp:
            if resp.status != 200:
                return None
            data = resp.read()
        if len(data) < MIN_BYTES or not data.startswith(b"\x89PNG"):
            return None
        out.write_bytes(data)
        return out
    except Exception as e:
        print(f"  failed: {domain}: {e}", file=sys.stderr)
        return None


def main(argv):
    if not argv:
        print(__doc__, file=sys.stderr)
        sys.exit(1)

    if argv[0] == "--file":
        with open(argv[1]) as f:
            domains = [line.strip() for line in f if line.strip() and not line.startswith("#")]
    else:
        domains = list(argv)

    ok, fail = 0, 0
    for d in domains:
        path = fetch_logo(d)
        if path:
            print(f"  ✓ {d} → {path.relative_to(REPO_ROOT)}")
            ok += 1
        else:
            print(f"  ✗ {d}", file=sys.stderr)
            fail += 1
    print(f"\nDone. {ok} downloaded, {fail} failed.")


if __name__ == "__main__":
    main(sys.argv[1:])
