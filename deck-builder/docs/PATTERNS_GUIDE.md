# Building the pattern library

The pattern library is one PowerPoint file, `patterns.pptx`, that holds your best human-made
slides with the client content replaced by `{{slots}}`. The builder copies these slides and fills
the slots, so the decks it makes look like your team made them. Quality here sets the quality of
every deck.

Keep this file on your PwC laptop. It never goes to GitHub.

## 1. Start from the brand template

1. Open the PwC brand template (.potx or the master .pptx) and save a copy as `patterns.pptx`.
2. Delete its sample slides. You now have the masters, layouts, fonts and colours with no slides.

Every pattern must live in this one file, built on this master. The builder copies slides within
the file, so layouts and images always resolve.

## 2. Collect 15-40 good slides

Pick slides people reuse: objectives (your 3-4 rotating styles), approach and timeline, team,
governance, why us, current vs future state, key messages, section dividers, title.

For each one: copy it from the old deck, then in `patterns.pptx` right-click → Paste Options →
**Use Destination Theme**. That puts it on the brand master. Fix anything that moved.

## 3. Clean each slide

- Remove client logos, client names and client photos.
- SmartArt: right-click → **Convert to Shapes**. The builder cannot fill SmartArt.
- Charts: leave them out for now. Version 1 fills text, icons and pictures, not chart data.
- Group each repeating item: select the box, its icon, head and body for item 1 → Ctrl+G. Repeat
  for each item. Groups let the builder remove item 4 cleanly when a deck only needs 3.

## 4. Add slots

You can do this by hand or let `prep_patterns.py` propose them (step 5).

By hand, replace each piece of text with a slot in double braces. Use these names so the builder
and the catalogue understand them:

| Text | Slot |
|---|---|
| Slide headline | `{{title}}` |
| Breadcrumb above the title | `{{tracker}}` |
| Source line | `{{source}}` |
| Item heads and bodies | `{{item1_head}}`, `{{item1_body}}`, `{{item2_head}}`, ... |
| Extra item fields | `{{item1_time}}`, `{{item1_owner}}`, `{{item1_metric}}` |
| Two-column slides | `{{left_head}}`, `{{left_body}}`, `{{right_head}}`, `{{right_body}}` |
| Anything else | a short lowercase name, letters, digits and underscores |

Rules:
- Type each slot in one go so PowerPoint keeps it in one piece. The builder handles split slots
  anyway, but typed-in-one-go is safer.
- The slot takes the formatting of its first character. Format the slot the way the text should look.
- For bullet lists, put the slot in one bulleted paragraph. A list in plan.json becomes one bullet per item.
- Leave numbers in badges ("1", "01", "A") as plain text. They are not slots.

Shape names (Home → Arrange → Selection Pane, or Alt+F10) are needed only for:
- **Icons**: name the picture `item1_icon`, `item2_icon`, ... Any picture works as a stand-in; the
  builder swaps the image and keeps the frame.
- **Item groups**: name the group `item1`, `item2`, ...
- **Item decoration** that is not grouped (a background box): `item1_bg`. Anything named `itemN...`
  is removed when item N is not used.

## 5. Let the script propose slots (optional)

Put the pasted, cleaned slides (still with their real text) in a file and run:

    python deck-builder/skill/pwc-deck-builder/scripts/prep_patterns.py my_slides.pptx --out private/

You get `my_slides_tokenised.pptx` (slots filled in, items and icons named) and
`my_slides_prep_report.md` (what it guessed, with the original text). Open the tokenised file,
press Alt+F10 and check every slide. The guesses are positional, so expect to fix some.

## 6. Describe each pattern in its speaker notes

The notes are how the agent knows when to use a pattern. Exactly this format:

    PATTERN: OBJ_PILLARS_3
    PURPOSE: Three objectives as pillars under one goal
    USE WHEN: Objectives with one overarching goal and 2-3 supporting outcomes
    ITEMS: 2-3
    AVOID: More than 15 words per pillar

- `PATTERN` IDs: uppercase, unique, describe the shape (`OBJ_3COL`, `TEAM_4`, `TIMELINE_GANTT_6`).
- Give each of your rotating objectives styles its own ID starting `OBJ_`, so the agent can rotate them.

## 7. Check it

    python deck-builder/skill/pwc-deck-builder/scripts/inventory.py private/patterns.pptx --catalogue

This prints the catalogue the agent will read: every pattern, its slots and a word capacity per
slot. If a slot is missing or a capacity looks wrong, fix the slide. Then build the zip
(docs/LAPTOP_SETUP.md).
