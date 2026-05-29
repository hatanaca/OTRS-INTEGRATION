#!/usr/bin/env python3
"""Valida mapa de visibilidade na fixture ticket 2840100."""
import re
import sys
from pathlib import Path

FIXTURE = Path(__file__).resolve().parents[1] / "tests/fixtures/ticket-2840100-article-table.html"
VISIBLE_IDS = {
    "14933276", "14933272", "14933269", "14933267", "14931479",
    "14931091", "14931078", "14931063",
}


def css_has_token(class_value: str, token: str) -> bool:
    return token in (class_value or "").split()


def row_visible_from_tr(row_tag: str) -> bool:
    m = re.search(r'class\s*=\s*"([^"]*)"', row_tag, re.I)
    if not m:
        return False
    c = m.group(1)
    if css_has_token(c, "NotVisibleForCustomer"):
        return False
    return css_has_token(c, "VisibleForCustomer")


def main() -> int:
    html = FIXTURE.read_text(encoding="utf-8")
    found = {}
    for m in re.finditer(
        r'(?s)<tr([^>]*\bid=["\']Row\d+["\'][^>]*)>.*?<input[^>]*class="ArticleID"[^>]*value="(\d+)"',
        html,
    ):
        found[m.group(2)] = row_visible_from_tr(m.group(1))
    exported = {aid for aid, v in found.items() if v}
    ok = exported == VISIBLE_IDS and len(found) == 14
    print(f"articles={len(found)} visible={len(exported)} ok={ok}")
    if not ok:
        print("expected visible:", sorted(VISIBLE_IDS))
        print("got visible:", sorted(exported))
        print("missing:", sorted(VISIBLE_IDS - exported))
        print("extra:", sorted(exported - VISIBLE_IDS))
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
