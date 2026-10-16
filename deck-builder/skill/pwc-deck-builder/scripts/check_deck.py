"""Quality checks on a built deck. Prints JSON and a readable summary.

    python check_deck.py deck.pptx [--plan plan.json] [--rules references/writing_rules.md]

Errors must be fixed before handing the deck over. Warnings need a look.
"""
import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402

# Figures that need a source: percentages, money, multiples, big numbers. Small counts ("12 interviews") are left alone.
NUMBER_RE = re.compile(r"(\d[\d,.]*\s?%|[$€£₹]\s?\d|\b\d[\d,.]*\s?(?:bn|mn|cr|crore|lakh|million|billion|x)\b"
                       r"|\b\d{1,3}(?:,\d{3})+\b|\b\d{3,}\b)", re.I)
YEAR_RE = re.compile(r"^(19|20)\d{2}$")
MAX_TITLE_WORDS = 15
MAX_EM_DASHES = 3


def load_banned(rules_path: Path) -> list[str]:
    """Phrases inside the ```banned fenced block of writing_rules.md, one per line."""
    if not rules_path.is_file():
        return []
    m = re.search(r"```banned\s*\n(.*?)```", rules_path.read_text(encoding="utf-8"), re.S)
    if not m:
        return []
    return [ln.strip() for ln in m.group(1).splitlines() if ln.strip() and not ln.strip().startswith("#")]


def is_title(shape) -> bool:
    if (shape.name or "").lower() == "title":
        return True
    return bool(getattr(shape, "is_placeholder", False) and shape.is_placeholder
                and shape.placeholder_format.type in (1, 3))


def has_real_numbers(text: str) -> bool:
    for m in NUMBER_RE.finditer(text):
        tok = m.group(0).strip()
        if YEAR_RE.match(tok) or re.fullmatch(r"0?\d", tok):
            continue
        return True
    return False


def check(deck_path, plan: dict | None = None, rules_path: Path | None = None) -> dict:
    prs = dk.open_deck(deck_path)
    rules_path = rules_path or dk.skill_dir() / "references" / "writing_rules.md"
    banned = load_banned(rules_path)
    banned_res = [(b, re.compile(r"(?<!\w)" + re.escape(b) + r"(?!\w)", re.I)) for b in banned]
    fonts_ok = dk.theme_fonts(prs)
    issues = []
    em_dashes = 0

    def add(n, level, msg):
        issues.append({"slide": n, "level": level, "issue": msg})

    for n, slide in enumerate(prs.slides, 1):
        slide_text, source_text = [], ""
        for shp, _ in dk.iter_shapes(slide.shapes):
            txt = dk.shape_text(shp)
            if not txt.strip():
                continue
            slide_text.append(txt)
            if (shp.name or "").lower().startswith("source") or txt.lower().startswith("source"):
                source_text = txt
            if "{{" in txt:
                add(n, "error", f"leftover token in '{shp.name}': {', '.join(dk.TOKEN_RE.findall(txt)) or txt[:40]}")
            if "[ICON:" in txt:
                add(n, "warning", f"icon to add by hand in '{shp.name}': {txt.strip()}")
            for tbc in re.findall(r"\[TBC[^\]]*\]", txt):
                add(n, "warning", f"open item in '{shp.name}': {tbc}")
            fit = dk.text_fit(shp)
            if fit and fit["needed_pt"] > fit["available_pt"] * 1.05:
                add(n, "error", f"text likely overflows '{shp.name}' (needs ~{fit['needed_pt']:.0f}pt, "
                                f"box {fit['available_pt']:.0f}pt). Shorten it.")
            if is_title(shp):
                words = len(txt.split())
                if words > MAX_TITLE_WORDS:
                    add(n, "warning", f"title has {words} words (max {MAX_TITLE_WORDS}): tighten it")
                elif words <= 3 and not re.search(r"\d", txt):
                    add(n, "warning", f"title '{txt.strip()}' is a label, not a message")
            for tf in dk.text_frames(shp):
                for p in tf.paragraphs:
                    for r in p.runs:
                        if r.font.name and not r.font.name.startswith("+") and fonts_ok and r.font.name not in fonts_ok:
                            add(n, "warning", f"off-template font '{r.font.name}' in '{shp.name}'")
                            break
        joined = "\n".join(slide_text)
        em_dashes += joined.count("—")
        for phrase, rx in banned_res:
            if rx.search(joined):
                add(n, "error", f"banned phrase: '{phrase}'")
        if has_real_numbers(joined) and not source_text.strip():
            add(n, "warning", "numbers on the slide but no source line")

    if em_dashes > MAX_EM_DASHES:
        add(0, "warning", f"{em_dashes} em dashes in the deck: rewrite most of them as two sentences or commas")

    if plan:
        pats = [str(s.get("pattern", "")).upper() for s in plan.get("slides", [])]
        for i in range(1, len(pats)):
            if pats[i] == pats[i - 1] and pats[i] not in ("SECTION",):
                add(i + 1, "warning", f"same pattern as previous slide ({pats[i]}): vary the layout")

    # dedupe font warnings per slide
    seen, unique = set(), []
    for it in issues:
        key = (it["slide"], it["issue"])
        if key not in seen:
            seen.add(key)
            unique.append(it)
    errors = sum(1 for i in unique if i["level"] == "error")
    return {"deck": str(deck_path), "slides": len(prs.slides), "errors": errors,
            "warnings": len(unique) - errors, "issues": unique}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--plan")
    ap.add_argument("--rules")
    a = ap.parse_args()
    plan = json.loads(Path(a.plan).read_text(encoding="utf-8")) if a.plan else None
    report = check(a.deck, plan, Path(a.rules) if a.rules else None)
    print(json.dumps(report, indent=2))
    print(f"\n{report['errors']} errors, {report['warnings']} warnings")
    for it in report["issues"]:
        where = f"slide {it['slide']}" if it["slide"] else "deck"
        print(f"  [{it['level']}] {where}: {it['issue']}")
    sys.exit(1 if report["errors"] else 0)


if __name__ == "__main__":
    main()
