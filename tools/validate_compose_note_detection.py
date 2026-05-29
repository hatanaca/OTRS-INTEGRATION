#!/usr/bin/env python3
"""Valida que AgentTicketZoom com links AgentTicketNote nao e confundido com popup de nota."""
from __future__ import annotations

import re
from pathlib import Path

FIXTURES = Path(__file__).resolve().parents[1] / "tests/fixtures"


def is_ticket_zoom(html: str) -> bool:
    if re.search(r'id="ArticleTable"|id="ArticleTableBody"', html, re.I):
        return True
    if re.search(r'<tr[^>]*\bid="Row\d+"[^>]*>[\s\S]*?class="ArticleID"', html, re.I | re.S):
        return True
    if re.search(r'WidgetSimple[^"\']*(?:VisibleForCustomer|NotVisibleForCustomer)', html, re.I):
        return True
    if re.search(r'ArticleIDs"\s*:\s*\[', html):
        return True
    if re.search(r'(?is)<a\s+[^>]*(?:id|name)\s*=\s*["\']Article\d+["\']', html):
        return True
    return False


def hidden_action(html: str) -> str:
    for pat in (
        r'(?is)<input[^>]*\bname\s*=\s*["\']Action["\'][^>]*\bvalue\s*=\s*["\']([^"\']+)["\']',
        r'(?is)<input[^>]*\bvalue\s*=\s*["\']([^"\']+)["\'][^>]*\bname\s*=\s*["\']Action["\']',
    ):
        m = re.search(pat, html)
        if m:
            return m.group(1).strip()
    return ""


def is_compose_note_form(html: str) -> bool:
    if is_ticket_zoom(html):
        return False
    if hidden_action(html) == "AgentTicketNote":
        return True
    if (
        re.search(r'id="Compose"', html)
        and re.search(r'name="IsVisibleForCustomer"', html)
        and not re.search(r'id="ArticleTable"', html, re.I)
    ):
        return True
    return False


def main() -> None:
    zoom = (FIXTURES / "ticket-2840100-zoom-with-note-link.html").read_text(encoding="utf-8")
    compose = (FIXTURES / "agent-ticket-note-compose.html").read_text(encoding="utf-8")

    assert "Action=AgentTicketNote" in zoom
    assert is_ticket_zoom(zoom), "fixture zoom deve ser AgentTicketZoom"
    assert not is_compose_note_form(zoom), "zoom com link de nota NAO deve ser compose"

    assert is_compose_note_form(compose), "popup AgentTicketNote deve ser detectado"
    assert not is_ticket_zoom(compose)

    print("OK: deteccao compose vs zoom validada")


if __name__ == "__main__":
    main()
