#!/usr/bin/env python3
"""Valida filtro DynamicField_Enviarpararelatorio=Sim na fixture AgentTicketZoom."""
from __future__ import annotations

import re
from pathlib import Path

FIXTURE = Path(__file__).resolve().parents[1] / "tests/fixtures/ticket-2840100-article-report-field.html"
FIELD_NAME = "Enviarpararelatorio"
EXPECTED = "Sim"


def article_rows(html: str) -> list[tuple[str, str]]:
    rows: list[tuple[str, str]] = []
    table = re.search(r'<table[^>]*id="ArticleTable"[^>]*>.*?</table>', html, re.I | re.S)
    scope = table.group(0) if table else html
    for m in re.finditer(r'<tr[^>]*id="Row\d+"[^>]*>(.*?)</tr>', scope, re.I | re.S):
        aid = re.search(r'class="ArticleID"[^>]*value="(\d+)"', m.group(1), re.I)
        if aid:
            rows.append((aid.group(1), m.group(0)))
    return rows


def article_scope(html: str, article_id: str) -> str:
    pat = rf'(?is)(<a\s+[^>]*(?:id|name)\s*=\s*["\']Article{re.escape(article_id)}["\'][^>]*>)(.*?)(?=<a\s+[^>]*(?:id|name)\s*=\s*["\']Article\d+["\']|\Z)'
    m = re.search(pat, html)
    return (m.group(1) + m.group(2)) if m else ""


def has_report_field(scope: str, field_name: str, expected: str) -> bool:
    fn = re.escape(field_name)
    val = re.escape(expected)
    patterns = [
        rf'DynamicField_{fn}[^>]*\bvalue\s*=\s*["\']{val}["\']',
        rf'name\s*=\s*["\']DynamicField_{fn}["\'][^>]*\bvalue\s*=\s*["\']{val}["\']',
        rf'<label[^>]*>\s*[^<]*Enviar[^<]*relat[^<]*</label>\s*<p[^>]*class\s*=\s*["\']Value["\'][^>]*>\s*{val}\s*</p>',
    ]
    return any(re.search(p, scope, re.I | re.S) for p in patterns)


def main() -> None:
    html = FIXTURE.read_text(encoding="utf-8")
    rows = article_rows(html)
    assert len(rows) == 2, rows
    report_ids = []
    for aid, _ in rows:
        scope = article_scope(html, aid)
        if has_report_field(scope, FIELD_NAME, EXPECTED):
            report_ids.append(aid)
    assert report_ids == ["14931078"], report_ids
    print(f"OK: {len(rows)} artigos, {len(report_ids)} com {FIELD_NAME}={EXPECTED} -> {report_ids}")


if __name__ == "__main__":
    main()
