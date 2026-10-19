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

## 5. Let the script add the slots

Keep the slides with their real text (the script records it as an example for each slot) and run,
from the repo folder in PowerShell:

    python deck-builder\skill\deck-pattern-builder\scripts\prep_patterns.py deck-builder\private\patterns_original.pptx --out deck-builder\private\prep

Always run it on the original slides, never on a tokenised copy. You get three files:

| File | What it is |
|---|---|
| `patterns_original_tokenised.pptx` | Text replaced by slots, table cells by `{{t1_r1c1}}`, items and icons named, speaker notes replaced by a PATTERN block with the original text as `EXAMPLE` lines |
| `patterns_original_prep_report.md` | Per slide: the slot it gave each piece of text, and anything it left alone. The top line lists the slides that need a manual look |
| `patterns_original_notes.csv` | One row per slide, to describe the patterns (step 6) |

What it handles and what it does not:

| On the slide | What happens |
|---|---|
| Text boxes, placeholders, grouped shapes | Slots (`title`, `tracker`, `source`, `itemN_head`, `itemN_body`, `textN`) |
| Tables | Every cell becomes `{{tN_rRcC}}` (empty cells too); the plan fills a table as a list of rows |
| Numbers and letters in badges ("1", "01", "A") | Left as they are on purpose |
| SmartArt | Left alone and listed. Right-click → Convert to Shapes, save, run the script again |
| Charts | Left alone and listed. Delete or replace with a table for now |
| Text inside pictures | Cannot be changed. Replace the picture or drop the slide |

Slot names on busy slides are positional (`item17_head`, `text9`). That is fine: the catalogue shows
the agent the original text of every slot, so it knows what each one is for. Very busy slides (more
than 8 items) are marked `complex layout: fill every slot`. Think about whether they are worth keeping.

## 6. Describe each pattern

Each pattern needs an ID and a description so the agent knows when to use it. With 100+ slides,
let the ChatPwC agent draft them:

1. In a chat with the Deck Builder agent, upload `..._tokenised.pptx` and `..._notes.csv` and say:
   "Describe the patterns in this CSV and apply them."
2. The agent fills `pattern`, `purpose`, `use_when`, `items`, `avoid` for every row (it writes `DROP` for
   duplicates), runs `apply_notes.py` and gives you back `patterns.pptx` and the filled CSV.
3. Open the CSV in Excel, fix what you disagree with (rename IDs, change DROP decisions), and apply it
   yourself:

       python deck-builder\skill\deck-pattern-builder\scripts\apply_notes.py deck-builder\private\prep\patterns_original_tokenised.pptx deck-builder\private\prep\patterns_original_notes.csv --out deck-builder\private\patterns.pptx

The notes it writes look like this (you can also edit them in PowerPoint):

    PATTERN: OBJ_PILLARS_3
    PURPOSE: Three objectives as pillars under one goal
    USE WHEN: Objectives with one overarching goal and 2-3 supporting outcomes
    ITEMS: 2-3
    AVOID: More than 15 words per pillar
    TABLE t1: 5 rows x 4 columns
    EXAMPLE title: Three outcomes will tell the client the programme worked

- `PATTERN` IDs: capitals, unique, describe the layout (`OBJ_3COL`, `TEAM_4`, `TIMELINE_GANTT_6`).
- Give each of your rotating objectives styles its own ID starting `OBJ_`.
- Keep the `EXAMPLE` and `TABLE` lines: the catalogue is built from them.

## 7. Check it

    python deck-builder/skill/deck-pattern-builder/scripts/inventory.py private/patterns.pptx --catalogue

This prints the catalogue the agent will read: every pattern, its slots and a word capacity per
slot. If a slot is missing or a capacity looks wrong, fix the slide. Then build the zip
(docs/LAPTOP_SETUP.md).
