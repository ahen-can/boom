---
name: pwc-deck-builder
description: Build PowerPoint proposal and client decks from RFPs, briefs, notes or other uploaded material, using the firm's template and a library of human-made slide patterns. Use when the user asks for a deck, presentation, slides, storyline or proposal in PowerPoint, or wants slides restyled into house patterns.
---

# PwC deck builder

You write the storyline and the words. Fixed scripts build the slides by copying human-made
pattern slides from `assets/patterns.pptx` and filling their `{{slots}}`. Never draw slides
yourself with python-pptx and never change the scripts. If something cannot be done with the
patterns, say so.

## Files

The skill folder is usually `/mnt/data/skills/pwc-deck-builder/`. If not, find it with
`find / -name build_deck.py -path "*pwc-deck-builder*" 2>/dev/null`. Below, `$SKILL` means that folder.

- `references/catalogue.md`: every pattern, what it is for, its slots and word capacity. Read before planning.
- `references/style_guide.md`: storyline flow, rhythm, density. Read before planning.
- `references/writing_rules.md`: how to write every headline and box, and the banned phrases. Read before writing.
- `references/storylines.md`: section orders and headlines from past human-made decks. Use as examples of flow and headline style only.
- `references/plan_schema.md`: the plan.json format.
- `scripts/`: `build_deck.py`, `check_deck.py`, `render.py`, `inventory.py`, `prep_patterns.py`, `extract_storylines.py`.

## Workflow

1. **Read the material.** Extract the client, what they are asking for, the evaluation criteria,
   any required structure, deadlines and every number worth using. Write a brief of 6-10 lines and
   list open questions.
2. **Propose the storyline. Then stop and wait.** Show a table with one row per slide:
   `# | section | headline | pattern | key content`. Headlines follow writing_rules.md. Patterns follow
   the rhythm rules in style_guide.md. Ask the user to approve or edit. Do not build before approval.
3. **Write the plan.** Save `/mnt/data/plan.json` following plan_schema.md, using only slot names from
   catalogue.md, within the word capacities. Validate it:
   `python $SKILL/scripts/build_deck.py /mnt/data/plan.json --validate-only`
4. **Build and check.**
   `python $SKILL/scripts/build_deck.py /mnt/data/plan.json`
   The result includes the check report. Fix every error by editing plan.json (shorten text, split a
   slide, change pattern) and rebuild. Look at the warnings and fix those that apply.
5. **Look at it.** `python $SKILL/scripts/render.py /mnt/data/<deck>.pptx` makes one PNG per slide and a
   contact sheet. If you can view images, check each slide for overflow, overlaps and empty areas, fix,
   and rebuild. Show the contact sheet to the user.
6. **Hand over.** Give the .pptx link from /mnt/data and a short list of what still needs a human:
   `[ICON: ...]` placeholders, warnings you left in place, facts you could not verify, open questions.

When the user asks for changes, edit plan.json and rebuild. Do not edit the .pptx directly.

## Hard rules

- Facts and numbers come only from the user's material. Never invent figures, client names,
  credentials, team members or case studies. Mark gaps as `[TBC: what is needed]` in the text.
- Every slide with figures gets a source line in its `source` slot.
- Follow writing_rules.md for every word, including speaker notes.
- One idea per slide. When content does not fit the capacity, split the slide or cut text.
- Do not use the same pattern on two consecutive slides (section dividers aside).

## Other jobs this skill can do

- "Which patterns do we have?": summarise catalogue.md.
- "Describe this deck": `python $SKILL/scripts/inventory.py <deck>` (shapes, text, positions) or
  `--layouts` (template layouts and placeholders).
- "Turn these slides into patterns": `python $SKILL/scripts/prep_patterns.py <deck> --out /mnt/data`.
  It writes `<name>_tokenised.pptx`, `<name>_prep_report.md` and `<name>_notes.csv`. Share all three and
  point out the slides the report lists under "needing a manual look".
- "Describe the patterns" (the user uploads `<name>_tokenised.pptx` and `<name>_notes.csv`):
  1. For each CSV row, read `slots` and `examples` (the human text each slot replaced). If you can
     view images, render the deck with render.py and look at each slide too.
  2. Fill `pattern` (a unique ID in capitals describing the layout: `OBJ_3COL_ICONS`, `GOV_MATRIX_3`,
     `TIMELINE_GANTT_6`; every objectives layout starts with `OBJ_`), `purpose` (one line, what the
     slide shows), `use_when` (when a deck needs it), `items` (for example `2-4`) and `avoid` (when not to
     use it). Put `DROP` as the pattern for slides that duplicate another pattern or are too
     client-specific to reuse. Work in batches of 20 rows and save the CSV after each batch.
  3. Run `python $SKILL/scripts/apply_notes.py <tokenised.pptx> <notes.csv> --out /mnt/data/patterns.pptx`,
     fix any errors it lists, then show `python $SKILL/scripts/inventory.py /mnt/data/patterns.pptx --catalogue`.
  4. Give the user `patterns.pptx` and the filled CSV to review.
- "Learn from these decks": `python $SKILL/scripts/extract_storylines.py <decks...> --out /mnt/data/storylines.md`
  and share it. The user reviews it and adds it to the skill.
