# Building the real skill on your PwC laptop

The repo holds the code and brand-neutral dummy assets. Your real pattern library, icons and
storylines stay on your laptop and only go into the zip you upload to ChatPwC.

## One-off setup

1. Download the repo: GitHub → `ahen-can/boom` → branch `claude/pwc-ppt-deck-builder-kqb3of` →
   Code → Download ZIP. Unzip it, for example to `C:\work\boom`.
2. Open the folder in VS Code. In the terminal:

       pip install python-pptx

3. Create a `private` folder inside `deck-builder` (it is git-ignored):

       deck-builder/
         private/
           patterns.pptx        your pattern library (docs/PATTERNS_GUIDE.md)
           icons/               PNG icons + icons_index.csv
           storylines.md        from extract_storylines.py
           style_guide.md       optional: your edited copy of references/style_guide.md
           writing_rules.md     optional: your edited copy of references/writing_rules.md
           past_decks/          decks to learn storylines from (input only, never zipped)

## Icons

1. Export the icons you use from the brand template: right-click an icon → Save as Picture → PNG.
   A transparent background and about 256 px square work well.
2. Name files plainly: `handshake.png`, `shield_risk.png`.
3. Create `icons_index.csv` in Excel and save as CSV:

       file,name,keywords
       handshake.png,handshake,partnership agreement collaboration trust
       shield_risk.png,shield,risk security protection compliance

The keywords are how the agent finds an icon from words like "risk management".

## Storylines

    python deck-builder/skill/pwc-deck-builder/scripts/extract_storylines.py --folder deck-builder/private/past_decks --out deck-builder/private/storylines.md

Open the result and delete anything confidential that should not sit in the skill. Headlines and
section orders are what matter.

## Build the zip

    python deck-builder/scripts/make_skill_zip.py --private deck-builder/private

It prints which of your files it used, the patterns it found and the zip size, and writes
`deck-builder/dist/pwc-deck-builder.zip`. It stops if the zip is over 50 MB (compress the pictures
in patterns.pptx with File → Compress Pictures).

## Test locally before uploading

    python deck-builder/skill/pwc-deck-builder/scripts/build_deck.py deck-builder/tests/sample_plan.json --patterns deck-builder/private/patterns.pptx --icons deck-builder/private/icons --out deck-builder/private/out

The sample plan uses the dummy pattern IDs, so it will list unknown patterns. Copy it, change the
pattern IDs and slot names to yours (see the catalogue), and run it again until it builds. Open the
result in PowerPoint.

## Upload to ChatPwC

1. Skills → find `pwc-deck-builder` → delete the old version (if any).
2. Upload skill → `deck-builder/dist/pwc-deck-builder.zip`.
3. Open the Deck Builder agent and check `pwc-deck-builder` is still in its skills list.

Repeat "Build the zip" and "Upload" every time you add patterns, icons or rules.
