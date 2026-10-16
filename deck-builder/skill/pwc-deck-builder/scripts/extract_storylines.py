"""Turn past human-made decks into storyline examples the agent can learn flow and headlines from.

    python extract_storylines.py deck1.pptx deck2.pptx ... [--folder DIR] [--out storylines.md] [--lines 2]

For each deck: section order, the headline of every slide, the layout, and the first lines of body
text. Headlines matter most: they show how the team writes "so what" titles.
Review the output and remove anything confidential before it goes into the skill.
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import deckkit as dk  # noqa: E402

EMU_PER_IN = 914400
SECTION_HINT = re.compile(r"section|divider|chapter|header", re.I)
LABEL_RE = re.compile(r"^\s*([0-9]{1,2}|[A-Z]|[ivx]{1,4})[.)]?\s*$", re.I)


def slide_title(slide, slide_h):
    for shp, _ in dk.iter_shapes(slide.shapes):
        if getattr(shp, "is_placeholder", False) and shp.is_placeholder \
                and shp.placeholder_format.type in (1, 3) and shp.has_text_frame and shp.text_frame.text.strip():
            return shp.text_frame.text.strip(), shp
    best, best_key = None, None
    for shp, _ in dk.iter_shapes(slide.shapes):
        if not (getattr(shp, "has_text_frame", False) and shp.has_text_frame) or not shp.text_frame.text.strip():
            continue
        if LABEL_RE.match(shp.text_frame.text):
            continue
        top = (shp.top or 0) / EMU_PER_IN
        p = shp.text_frame.paragraphs[0]
        size = dk.effective_font_pt(shp, p.runs[0] if p.runs else None, p)
        key = (top < slide_h * 0.3, size, -top)
        if best_key is None or key > best_key:
            best, best_key = shp, key
    return (best.text_frame.text.strip(), best) if best else ("", None)


def body_lines(slide, title_shape, n):
    out = []
    title_el = title_shape._element if title_shape is not None else None
    title_top = (title_shape.top or 0) if title_shape is not None else 0
    for shp, parents in dk.iter_shapes(slide.shapes):
        if shp._element is title_el:
            continue
        if not parents and title_el is not None and (shp.top or 0) + (shp.height or 0) <= title_top:
            continue  # tracker / breadcrumb above the title
        for tf in dk.text_frames(shp):
            for p in tf.paragraphs:
                t = "".join(r.text for r in p.runs).strip()
                if len(t) > 3 and not re.match(r"^(source|note)s?\b", t, re.I):
                    out.append(t)
    return out[:n]


def is_section(slide, title, word_count):
    return bool(SECTION_HINT.search(slide.slide_layout.name)) or (title and word_count <= 8)


def deck_story(path: Path, n_lines: int) -> str:
    prs = dk.open_deck(path)
    slide_h = prs.slide_height / EMU_PER_IN
    out = [f"## {path.stem}", ""]
    sections = []
    for i, slide in enumerate(prs.slides, 1):
        title, tshape = slide_title(slide, slide_h)
        title = " ".join(title.split())
        words = sum(len(dk.shape_text(s).split()) for s, _ in dk.iter_shapes(slide.shapes))
        layout = slide.slide_layout.name
        if is_section(slide, title, words):
            sections.append(title or f"(slide {i})")
            out.append(f"**[Section] {title}**  _(slide {i}, {layout})_")
            continue
        out.append(f"{i}. {title or '(no title)'}  _({layout}, {words} words)_")
        for line in body_lines(slide, tshape, n_lines):
            out.append(f"    - {line[:140]}")
    out.insert(2, f"Sections: {' → '.join(sections) if sections else 'none detected'}")
    out.insert(3, "")
    return "\n".join(out) + "\n"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("decks", nargs="*")
    ap.add_argument("--folder", help="read every .pptx in this folder")
    ap.add_argument("--out", default="storylines.md")
    ap.add_argument("--lines", type=int, default=2, help="body lines to keep per slide (0 = headlines only)")
    a = ap.parse_args()
    paths = [Path(p) for p in a.decks]
    if a.folder:
        paths += sorted(Path(a.folder).glob("*.pptx"))
    paths = [p for p in paths if not p.name.startswith("~$")]
    if not paths:
        ap.error("give decks or --folder")
    parts = ["# Storyline examples", "",
             "Extracted from past human-made decks. Use them for section order, slide count per section and",
             "headline style. Do not copy client facts from here into a new deck.", ""]
    for p in paths:
        try:
            parts.append(deck_story(p, a.lines))
        except Exception as e:  # noqa: BLE001
            parts.append(f"## {p.stem}\n\n(could not read: {type(e).__name__}: {e})\n")
    Path(a.out).write_text("\n".join(parts), encoding="utf-8")
    print(f"wrote {a.out} from {len(paths)} deck(s)")


if __name__ == "__main__":
    main()
