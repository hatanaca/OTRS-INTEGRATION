#!/usr/bin/env python3
"""Falha se .ps1 tiver travessao Unicode ou estiver sem UTF-8 BOM (Menu-OTRS.ps1)."""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "Menu-OTRS.ps1"
BAD = ("\u2014", "\u2013", "\u201c", "\u201d", "\u2018", "\u2019")
errors = []


def check_file(path: Path) -> None:
    raw = path.read_bytes()
    if path.name == "Menu-OTRS.ps1":
        if not raw.startswith(b"\xef\xbb\xbf"):
            errors.append(f"{path}: falta UTF-8 BOM (necessario para Windows PowerShell 5.1)")
        while raw.startswith(b"\xef\xbb\xbf"):
            raw = raw[3:]
        if raw.startswith(b"\xef\xbb\xbf"):
            errors.append(f"{path}: BOM duplicado")
    text = raw.decode("utf-8")
    for i, line in enumerate(text.splitlines(), 1):
        for ch in line:
            if ch in BAD or (0x2010 <= ord(ch) <= 0x201F):
                errors.append(f"{path}:{i}: caractere U+{ord(ch):04X} ({ch!r})")
                break


def main() -> int:
    check_file(MAIN)
    for ps in (ROOT / "scripts").glob("*.ps1"):
        check_file(ps)
    if errors:
        for e in errors:
            print(e, file=sys.stderr)
        return 1
    print("encoding OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
