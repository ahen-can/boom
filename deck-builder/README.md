# Deck builder for ChatPwC

A ChatPwC skill that turns RFPs, briefs and notes into draft PowerPoint decks. The model writes
the storyline and the words; fixed Python scripts build the slides by copying human-made pattern
slides and filling their `{{slots}}`, so the result looks like your team's work.

## What is here

| Path | What it is |
|---|---|
| `dist/deck-pattern-builder.zip` | The skill, ready to upload (dummy brand-neutral patterns and icons) |
| `skill/deck-pattern-builder/` | The skill source: `SKILL.md`, `scripts/`, `references/`, `assets/` |
| `agent/instructions.md` | What to paste into the ChatPwC agent builder |
| `docs/PATTERNS_GUIDE.md` | How to turn your best slides into the pattern library |
| `docs/LAPTOP_SETUP.md` | How to build the real zip on your laptop with your private files |
| `tests/` | End-to-end tests and a sample plan |
| `probe-skill/`, `dist/pptx-probe*.zip` | Step 0: the sandbox probe (passed) |

## How a deck gets made

1. Upload the RFP or notes to the Deck Builder agent.
2. The agent writes a brief and a storyline table (headline + pattern per slide) and waits.
3. You approve or edit the storyline.
4. The agent writes `plan.json` and runs `build_deck.py`, which copies the chosen pattern slides,
   fills the slots, removes unused items, swaps icons, then runs `check_deck.py`.
5. `check_deck.py` fails the deck on leftover slots, text overflow and banned phrases, and warns on
   label-style titles, missing sources, repeated patterns and off-template fonts. The agent fixes the plan
   and rebuilds.
6. `render.py` makes slide images and a contact sheet; you get the .pptx.

## Getting started

1. Upload `dist/deck-pattern-builder.zip` in ChatPwC (Skills → Upload skill).
2. Create the agent from `agent/instructions.md`.
3. Try "Build a proposal deck from this RFP" with a test document. The dummy patterns are
   TITLE, SECTION, OBJ_3COL, OBJ_NUMBERED_4, APPROACH_CHEVRON_4, TWO_COL and GOV_MATRIX.
4. Build your real pattern library (`docs/PATTERNS_GUIDE.md`), then the real zip
   (`docs/LAPTOP_SETUP.md`), and upload that instead.

## Scripts

All in `skill/deck-pattern-builder/scripts/`, Python 3.10+ with python-pptx.

| Script | Use |
|---|---|
| `build_deck.py plan.json` | Build a deck (add `--validate-only` to check the plan) |
| `check_deck.py deck.pptx` | Quality checks |
| `render.py deck.pptx` | PNG per slide and a contact sheet (needs LibreOffice) |
| `inventory.py deck.pptx` | Shapes, names, positions, text; `--layouts` for template layouts; `--catalogue` for the pattern catalogue |
| `prep_patterns.py slides.pptx` | Add slots (text, tables), name items and icons, record example text, write a notes CSV |
| `apply_notes.py deck.pptx notes.csv` | Write pattern IDs and descriptions from the CSV into the notes (`DROP` removes a slide) |
| `extract_storylines.py --folder decks/` | Section order and headlines from past decks |

Tests: `pip install pytest && python -m pytest deck-builder/tests -q`

Rebuild the dummy assets: `python deck-builder/scripts/make_dummy_assets.py`, then
`python deck-builder/scripts/make_skill_zip.py`.

Credits and licences: `THIRD_PARTY_NOTICES.md`.
