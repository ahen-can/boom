"""Build a deck from a slide plan (plan.json) and the pattern library.

    python build_deck.py plan.json [--patterns assets/patterns.pptx] [--icons assets/icons] [--out DIR]
    python build_deck.py plan.json --validate-only

See references/plan_schema.md for the plan format. The output deck is a copy of the pattern
library: chosen pattern slides are duplicated in plan order, filled, and the originals removed,
so masters, fonts and theme come from the template untouched.
"""
import argparse
import json
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402
from check_deck import check  # noqa: E402

ICON_KEY_RE = re.compile(r"icon", re.I)


def load_plan(path: Path) -> dict:
    plan = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(plan.get("slides"), list) or not plan["slides"]:
        raise SystemExit("plan.json needs a non-empty 'slides' list")
    return plan


def validate(plan: dict, patterns: dict) -> tuple[list[str], list[str]]:
    errors, warnings = [], []
    defaults = plan.get("defaults", {})
    for n, s in enumerate(plan["slides"], 1):
        pid = str(s.get("pattern", "")).upper()
        if pid not in patterns:
            errors.append(f"slide {n}: unknown pattern '{s.get('pattern')}'. Available: {', '.join(sorted(patterns))}")
            continue
        _, pslide = patterns[pid]
        tokens = set(dk.slide_tokens(pslide))
        icon_shapes = {shp.name for shp, _ in dk.iter_shapes(pslide.shapes) if ICON_KEY_RE.search(shp.name or "")}
        fill = s.get("fill", {})
        tables = dk.table_dims(pslide)
        plain = sorted(t for t in tokens if not dk.TABLE_TOKEN_RE.match(t))
        unknown = [k for k in fill if k not in tokens and k not in tables]
        if unknown:
            errors.append(f"slide {n} ({pid}): unknown slots {unknown}. This pattern has: {plain}"
                          + (f" and tables {sorted(tables)}" if tables else ""))
        for t, (rows, cols) in tables.items():
            v = fill.get(t)
            if v is None:
                continue
            if not isinstance(v, list) or not all(isinstance(r, list) for r in v):
                errors.append(f"slide {n} ({pid}): table '{t}' must be a list of rows, each a list of cells")
                continue
            if len(v) > rows:
                errors.append(f"slide {n} ({pid}): table '{t}' has {rows} rows, plan gives {len(v)}. "
                              "Split the content over two slides or merge rows.")
            if any(len(r) > cols for r in v):
                errors.append(f"slide {n} ({pid}): table '{t}' has {cols} columns, a row in the plan has more")
        bad_icons = [k for k in s.get("icons", {}) if k not in icon_shapes]
        if bad_icons:
            errors.append(f"slide {n} ({pid}): no icon shapes named {bad_icons}. Icon shapes: {sorted(icon_shapes)}")
        if "title" in tokens and not (fill.get("title") or defaults.get("title")):
            warnings.append(f"slide {n} ({pid}): no title")
        for k, v in fill.items():
            if k in tables:
                continue
            if not isinstance(v, (str, list, int, float)) and v is not None:
                errors.append(f"slide {n}: slot '{k}' must be text or a list of text")
    return errors, warnings


def build(plan: dict, patterns_path: Path, icons_dir: Path, out_dir: Path) -> dict:
    prs = dk.open_deck(patterns_path)
    patterns = dk.index_patterns(prs)
    errors, warnings = validate(plan, patterns)
    if errors:
        return {"ok": False, "errors": errors, "warnings": warnings}

    originals = [s for s in prs.slides]
    icon_index = dk.load_icon_index(icons_dir) if icons_dir.is_dir() else []
    defaults = plan.get("defaults", {})
    log = []

    new_slides = [dk.duplicate_slide(prs, patterns[str(s["pattern"]).upper()][1]) for s in plan["slides"]]
    for s in originals:
        dk.delete_slide(prs, s)

    for n, (spec, slide) in enumerate(zip(plan["slides"], new_slides), 1):
        values = dict(defaults)
        if spec.get("section"):
            values["tracker"] = spec["section"]
        values.update(spec.get("fill", {}))
        filled, unfilled = dk.fill_tokens(slide, values)
        _, rows_given = dk.expand_tables(values)
        for note in dk.prune_unused(slide, filled, unfilled, rows_given):
            log.append(f"slide {n}: {note}")
        for shape_name, query in spec.get("icons", {}).items():
            png = dk.find_icon(query, icons_dir, icon_index) if icon_index else None
            log.append(f"slide {n}: " + dk.set_icon(slide, shape_name, png, query))
        for shp, _ in dk.iter_shapes(slide.shapes):
            note = dk.fit_text(shp)
            if note:
                log.append(f"slide {n}: {note}")
        dk.set_notes(slide, spec.get("notes", ""))

    out_dir.mkdir(parents=True, exist_ok=True)
    name = plan.get("output") or "deck.pptx"
    if not name.lower().endswith(".pptx"):
        name += ".pptx"
    out = out_dir / Path(name).name
    prs.save(str(out))
    report = check(out, plan=plan)
    return {"ok": True, "output": str(out), "slides": len(new_slides), "warnings": warnings, "log": log,
            "check": report}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("plan")
    ap.add_argument("--patterns", default=str(dk.skill_dir() / "assets" / "patterns.pptx"))
    ap.add_argument("--icons", default=str(dk.skill_dir() / "assets" / "icons"))
    ap.add_argument("--out", default=None, help="output folder (default /mnt/data, else current folder)")
    ap.add_argument("--validate-only", action="store_true")
    a = ap.parse_args()

    plan = load_plan(Path(a.plan))
    if a.validate_only:
        prs = dk.open_deck(a.patterns)
        errors, warnings = validate(plan, dk.index_patterns(prs))
        print(json.dumps({"ok": not errors, "errors": errors, "warnings": warnings}, indent=2))
        sys.exit(1 if errors else 0)

    result = build(plan, Path(a.patterns), Path(a.icons), Path(a.out) if a.out else dk.default_out_dir())
    print(json.dumps(result, indent=2))
    sys.exit(0 if result["ok"] else 1)


if __name__ == "__main__":
    main()
