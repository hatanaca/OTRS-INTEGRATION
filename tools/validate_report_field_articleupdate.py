#!/usr/bin/env python3
"""Valida deteccao Enviar para relatorio quando ArticleMetaFields so vem via ArticleUpdate."""
from __future__ import annotations

import re
from pathlib import Path

FIXTURES = Path(__file__).resolve().parents[1] / "tests/fixtures"
FIELD_NAME = "Enviarpararelatorio"
EXPECTED = "Sim"


def article_rows(html: str) -> list[tuple[str, str, str]]:
    rows: list[tuple[str, str, str]] = []
    table = re.search(r'<table[^>]*id="ArticleTable"[^>]*>.*?</table>', html, re.I | re.S)
    scope = table.group(0) if table else html
    for m in re.finditer(r'<tr([^>]*id="Row\d+"[^>]*)>(.*?)</tr>', scope, re.I | re.S):
        tag, body = m.group(1), m.group(2)
        aid = re.search(r'class="ArticleID"[^>]*value="(\d+)"', body, re.I)
        if aid:
            rows.append((aid.group(1), tag, body))
    return rows


def article_count(row_tag: str, row_html: str) -> int:
    src = row_tag + row_html
    m = re.search(r"Subaction=ArticleUpdate;Count=(\d+)", src)
    if m:
        return int(m.group(1))
    m = re.search(r'id="Row(\d+)"', row_tag)
    if m:
        return int(m.group(1))
    return 0


def report_value(article_html: str, field_name: str) -> str:
    fn = re.escape(field_name)
    pats = [
        rf"<label[^>]*>\s*[^<]*Enviar[^<]*relat[^<]*</label>\s*<(?:p|span)[^>]*class\s*=\s*['\"]Value['\"][^>]*>\s*([^<]+)\s*</(?:p|span)>",
        rf"name\s*=\s*['\"]DynamicField_{fn}['\"][^>]*\bvalue\s*=\s*['\"]([^'\"]+)['\"]",
    ]
    for p in pats:
        m = re.search(p, article_html, re.I | re.S)
        if m:
            return m.group(1).strip()
    return ""


def value_matches(actual: str, expected: str) -> bool:
    a = actual.strip().lower()
    e = expected.strip().lower()
    if a == e:
        return True
    if e == "sim" and a in ("sim", "1", "yes", "true"):
        return True
    return False


def main() -> None:
    zoom = (FIXTURES / "ticket-2840100-table-only-empty-widgets.html").read_text(encoding="utf-8")
    update = (FIXTURES / "agent-ticket-article-update-report-sim.html").read_text(encoding="utf-8")

    rows = article_rows(zoom)
    assert len(rows) == 1, rows
    aid, tag, body = rows[0]
    assert aid == "14931078"
    assert article_count(tag, body) == 2

    # Widget vazio na pagina principal
    scope_pat = rf'(?is)<a[^>]*Article{re.escape(aid)}[^>]*>.*?(?=<a[^>]*Article\d+|\Z)'
    scope = re.search(scope_pat, zoom)
    assert scope, "ancora do artigo"
    assert not value_matches(report_value(scope.group(0), FIELD_NAME), EXPECTED)

    # ArticleUpdate traz ArticleMetaFields
    assert value_matches(report_value(update, FIELD_NAME), EXPECTED), report_value(update, FIELD_NAME)

    print("OK: tabela sem meta + ArticleUpdate com Sim validado")


if __name__ == "__main__":
    main()
