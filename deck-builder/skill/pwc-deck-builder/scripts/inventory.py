"""Describe a deck so a model (or a person) can work with it without opening PowerPoint.

    python inventory.py deck.pptx                 # every slide: shapes, names, positions, font sizes, text, tokens
    python inventory.py deck.pptx --layouts       # slide size, layouts, placeholders, content area
    python inventory.py patterns.pptx --catalogue [--write references/catalogue.md]

--layouts follows the template-analysis approach of Anthropic's ppt-template-creator skill
(Apache-2.0, github.com/anthropics/financial-services): list layouts with their placeholder
idx/type/position so content can be placed inside the template's real content area.
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402

EMU_PER_IN = 914400


def inches(v):
    return round(v / EMU_PER_IN, 2) if v is not None else None


def shape_info(shp, parents):
    info = {
        "name": shp.name,
        "type": str(shp.shape_type).split(".")[-1].split(" ")[0] if shp.shape_type else "UNKNOWN",
        "group": "/".join(p.name for p in parents) or None,
        "x": inches(shp.left), "y": inches(shp.top), "w": inches(shp.width), "h": inches(shp.height),
    }
    if getattr(shp, "is_placeholder", False) and shp.is_placeholder:
        info["placeholder"] = {"idx": shp.placeholder_format.idx,
                               "type": str(shp.placeholder_format.type).split(".")[-1].split(" ")[0]}
    txt = dk.shape_text(shp)
    if txt.strip():
        info["text"] = txt if len(txt) <= 300 else txt[:300] + "…"
        tf = shp.text_frame if getattr(shp, "has_text_frame", False) and shp.has_text_frame else None
        if tf and tf.paragraphs:
            p = tf.paragraphs[0]
            info["font_pt"] = dk.effective_font_pt(shp, p.runs[0] if p.runs else None, p)
            info["bold"] = bool(p.runs and p.runs[0].font.bold)
        toks = dk.tokens_in(shp)
        if toks:
            info["tokens"] = toks
    return info


def slides_inventory(prs):
    out = []
    for n, s in enumerate(prs.slides, 1):
        out.append({
            "slide": n,
            "layout": s.slide_layout.name,
            "pattern": dk.pattern_id(s),
            "notes": dk.notes_text(s)[:500],
            "shapes": [shape_info(shp, parents) for shp, parents in dk.iter_shapes(s.shapes)],
        })
    return out


def layouts_inventory(prs):
    layouts = []
    for i, layout in enumerate(prs.slide_layouts):
        phs = []
        for ph in layout.placeholders:
            phs.append({"idx": ph.placeholder_format.idx,
                        "type": str(ph.placeholder_format.type).split(".")[-1].split(" ")[0],
                        "name": ph.name, "x": inches(ph.left), "y": inches(ph.top),
                        "w": inches(ph.width), "h": inches(ph.height)})
        layouts.append({"index": i, "name": layout.name, "placeholders": phs})
    # Content area: the body/object placeholders on content layouts, else between title and footer.
    bodies = [p for lay in layouts for p in lay["placeholders"] if p["type"] in ("BODY", "OBJECT")]
    area = None
    if bodies:
        b = max(bodies, key=lambda p: (p["w"] or 0) * (p["h"] or 0))
        area = {"x": b["x"], "y": b["y"], "w": b["w"], "h": b["h"], "from": "largest body/object placeholder"}
    return {"slide_width_in": inches(prs.slide_width), "slide_height_in": inches(prs.slide_height),
            "theme_fonts": sorted(dk.theme_fonts(prs)), "content_area": area, "layouts": layouts}


def catalogue(prs) -> str:
    lines = [
        "# Pattern catalogue",
        "",
        "Generated from assets/patterns.pptx by `inventory.py --catalogue`. Do not edit by hand: change the",
        "pattern slide's speaker notes (PATTERN / PURPOSE / USE WHEN / ITEMS) and regenerate.",
        "",
        "Slots: fill these keys in plan.json. `~N words` is a rough capacity at the template font size.",
        "Item slots (`itemN_...`) can be left out from the end: unused items are removed with their shapes.",
        "",
    ]
    for s in prs.slides:
        pid = dk.pattern_id(s)
        if not pid:
            continue
        meta = dk.pattern_meta(s)
        lines.append(f"## {pid}")
        for key in ("PURPOSE", "USE WHEN", "ITEMS", "AVOID"):
            if meta.get(key):
                lines.append(f"- **{key.title()}:** {meta[key]}")
        slots, icons = [], []
        for shp, _ in dk.iter_shapes(s.shapes):
            for t in dk.tokens_in(shp):
                cap = dk.capacity_words(shp) if getattr(shp, "has_text_frame", False) and shp.has_text_frame else None
                slots.append(f"`{t}`" + (f" (~{cap} words)" if cap else ""))
            if "icon" in (shp.name or "").lower():
                icons.append(f"`{shp.name}`")
        lines.append(f"- **Slots:** {', '.join(slots) if slots else 'none'}")
        if icons:
            lines.append(f"- **Icons:** {', '.join(icons)}")
        lines.append("")
    return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("deck")
    ap.add_argument("--layouts", action="store_true")
    ap.add_argument("--catalogue", action="store_true")
    ap.add_argument("--write", help="write output to this file instead of printing")
    a = ap.parse_args()
    prs = dk.open_deck(a.deck)
    if a.catalogue:
        text = catalogue(prs)
    elif a.layouts:
        text = json.dumps(layouts_inventory(prs), indent=2)
    else:
        text = json.dumps(slides_inventory(prs), indent=2, ensure_ascii=False)
    if a.write:
        Path(a.write).write_text(text + "\n", encoding="utf-8")
        print(f"wrote {a.write}")
    else:
        print(text)


if __name__ == "__main__":
    main()
