#!/usr/bin/env python3
"""Valida contagem VisibleForCustomer na fixture AgentTicketZoom (ticket 2840100)."""
import re
import sys
from pathlib import Path

FIXTURE = Path(__file__).resolve().parents[1] / "tests/fixtures/ticket-2840100-article-table.html"
EXPECTED_VISIBLE = 8
EXPECTED_NOT_VISIBLE = 6


def css_has_token(class_value: str, token: str) -> bool:
    return token in (class_value or "").split()


def row_visible(row_tag: str) -> bool:
    m = re.search(r'class\s*=\s*"([^"]*)"', row_tag, re.I)
    if not m:
        return False
    classes = m.group(1)
    if css_has_token(classes, "NotVisibleForCustomer"):
        return False
    return css_has_token(classes, "VisibleForCustomer")


def main() -> int:
    html = FIXTURE.read_text(encoding="utf-8")
    rows = re.findall(r"(?s)<tr([^>]*\bid=[\"']Row\d+[\"'][^>]*)>", html)
    visible = sum(1 for r in rows if row_visible(r))
    not_visible = sum(
        1
        for r in rows
        if (m := re.search(r'class\s*=\s*"([^"]*)"', r, re.I))
        and css_has_token(m.group(1), "NotVisibleForCustomer")
    )
    ok = visible == EXPECTED_VISIBLE and not_visible == EXPECTED_NOT_VISIBLE
    print(f"rows={len(rows)} visible={visible} not_visible={not_visible} ok={ok}")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
