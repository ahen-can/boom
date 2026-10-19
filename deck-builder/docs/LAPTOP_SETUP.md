# Building and uploading the real skill (PwC laptop)

The repo holds the code and dummy assets. Your pattern library, brand icons and storylines stay on
your laptop and only go into the zip you upload to ChatPwC.

## 1. Get the repo in one piece

The scripts find each other by folder, so keep the repo's layout. Don't copy the `.py` files out
into another folder.

1. GitHub → `ahen-can/boom` → branch `claude/pwc-ppt-deck-builder-kqb3of` → Code → Download ZIP.
2. Unzip to `Desktop\repo\boom` (any folder works; the commands below assume you are in it).
3. Once: `pip install python-pptx`

## 2. Put your files in `deck-builder\private\`

Create the folder (git ignores it) and copy in:

| File | What it is |
|---|---|
| `patterns.pptx` | The finished library: the `patterns.pptx` ChatPwC gave you after applying the notes |
| `brand_icons.pptx` | Only the 4-5 PwC icon and pictogram slides (copy them into a new file from the brand template) |
| `past_decks\` | Optional: reference decks to learn storylines from |

## 3. Shrink the pattern library (recommended)

Open `patterns.pptx` → select any picture → Picture Format → Compress Pictures → untick "Apply only to
this picture" → 150 ppi → OK → Save. 30+ MB decks are slow to build and render in ChatPwC; this
usually brings them under 10 MB with no visible change.

## 4. Index the brand icons

From `Desktop\repo\boom` in PowerShell:

    python deck-builder\skill\deck-pattern-builder\scripts\index_icons.py deck-builder\private\brand_icons.pptx --out deck-builder\private

This writes `icons.pptx` (each icon named `icon:<name>`, check with Alt+F10), `icons.csv` and `icons.md`.
Icons take their name from the text label under or beside them, so labelled icon slides give the
best names. Unlabelled ones become `icon_s<slide>_<n>`: rename them in `icons.csv` if you like.

Then make the icons findable: open `icons.csv` in Excel and add synonyms to the `keywords` column, or
ask the agent ("Add keywords to the icons", attaching `icons.csv` and `icons.pptx` to the code
environment) and save the CSV it returns over yours. Keep `slide` and `shape_id` as they are.

## 5. Storylines (optional)

    python deck-builder\skill\deck-pattern-builder\scripts\extract_storylines.py --folder deck-builder\private\past_decks --out deck-builder\private\storylines.md

Read it and delete anything confidential.

## 6. Build the zip

    python deck-builder\scripts\make_skill_zip.py --private deck-builder\private

It prints the files it used, how many patterns and icons it found and the zip size, and writes
`deck-builder\dist\deck-pattern-builder.zip`. It stops if the zip is over 50 MB (go back to step 3).

## 7. Upload

1. ChatPwC → Skills → delete the old `deck-pattern-builder`.
2. Upload skill → `deck-builder\dist\deck-pattern-builder.zip`.
3. Open your agent → Skills → make sure `deck-pattern-builder` is selected. File Search off.

Repeat steps 6 and 7 whenever you change patterns, icons, storylines or rules.

## Optional: test on your laptop first

    python deck-builder\skill\deck-pattern-builder\scripts\inventory.py deck-builder\private\patterns.pptx --catalogue

shows what the agent will see. To build a test deck locally, copy `deck-builder\tests\sample_plan.json`,
change the pattern IDs and slots to yours, and run:

    python deck-builder\skill\deck-pattern-builder\scripts\build_deck.py my_plan.json --patterns deck-builder\private\patterns.pptx --icons-pptx deck-builder\private\icons.pptx --out deck-builder\private\out
