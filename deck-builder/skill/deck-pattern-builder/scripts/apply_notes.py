"""Write pattern descriptions from a CSV into each slide's speaker notes.

    python apply_notes.py patterns_tokenised.pptx patterns_tokenised_notes.csv [--out patterns.pptx]

CSV columns (from prep_patterns.py): slide, pattern, purpose, use_when, items, avoid. Other columns
are ignored. Put DROP in the pattern column to delete that slide from the library. The TABLE and
EXAMPLE lines already in the notes are kept.
"""
import argparse
import csv
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402

ID_RE = re.compile(r"^[A-Z][A-Z0-9_]*$")
MONTHS = {m: i for i, m in enumerate("jan feb mar apr may jun jul aug sep oct nov dec".split(), 1)}


def fix_excel_range(value: str) -> str:
    """Excel turns item ranges like 1-8 into dates (01-Aug, Aug-01, 1/8/2026). Turn them back."""
    v = value.strip()
    m = re.fullmatch(r"(\d{1,2})-([A-Za-z]{3})", v) or re.fullmatch(r"([A-Za-z]{3})-(\d{1,2})", v)
    if m:
        a, b = m.groups()
        day, mon = (a, b) if a.isdigit() else (b, a)
        if mon.lower() in MONTHS:
            return f"{int(day)} to {MONTHS[mon.lower()]}"
    m = re.fullmatch(r"(\d{1,2})/(\d{1,2})/\d{2,4}", v)
    if m:
        return f"{int(m.group(1))} to {int(m.group(2))}"
    return v


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("csv")
    ap.add_argument("--out", help="output deck (default: <deck>_described.pptx)")
    a = ap.parse_args()
    deck = Path(a.deck)
    prs = dk.open_deck(deck)
    with open(a.csv, newline="", encoding="utf-8-sig") as f:
        rows = [{(k or "").strip().lower(): (v or "").strip() for k, v in r.items()} for r in csv.DictReader(f)]

    slides = list(prs.slides)
    errors, warnings, seen, drop = [], [], {}, []
    for r in rows:
        try:
            n = int(r.get("slide", ""))
        except ValueError:
            errors.append(f"bad slide number: {r.get('slide')!r}")
            continue
        if not 1 <= n <= len(slides):
            errors.append(f"slide {n} does not exist (deck has {len(slides)})")
            continue
        pid = r.get("pattern", "").upper().replace(" ", "_").replace("-", "_")
        if pid == "DROP":
            drop.append(n)
            continue
        if not ID_RE.match(pid):
            errors.append(f"slide {n}: pattern '{r.get('pattern')}' must be letters, digits and _ (e.g. OBJ_3COL)")
            continue
        if pid in seen:
            errors.append(f"pattern {pid} used on slides {seen[pid]} and {n}: IDs must be unique")
            continue
        seen[pid] = n
        if not r.get("purpose") or not r.get("use_when"):
            warnings.append(f"slide {n} ({pid}): purpose or use_when is empty")
        slide = slides[n - 1]
        keep = [ln for ln in dk.notes_text(slide).splitlines()
                if dk.EXAMPLE_RE.match(ln) or dk.TABLE_NOTE_RE.match(ln)]
        head = [f"PATTERN: {pid}", f"PURPOSE: {r.get('purpose', '')}", f"USE WHEN: {r.get('use_when', '')}",
                f"ITEMS: {fix_excel_range(r.get('items', ''))}"]
        if r.get("avoid"):
            head.append(f"AVOID: {r['avoid']}")
        dk.set_notes(slide, "\n".join(head + keep))

    if errors:
        print("Nothing written. Fix these in the CSV:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)
    for n in sorted(drop, reverse=True):
        dk.delete_slide(prs, slides[n - 1])
    missing = [i + 1 for i, s in enumerate(slides) if i + 1 not in drop and i + 1 not in seen.values()]
    out = Path(a.out) if a.out else deck.with_name(f"{deck.stem}_described.pptx")
    prs.save(str(out))
    print(f"wrote {out}: {len(seen)} patterns described, {len(drop)} slides dropped")
    if missing:
        print(f"  slides not in the CSV (kept with their current notes): {missing}")
    for w in warnings:
        print(f"  warning: {w}")


if __name__ == "__main__":
    main()
