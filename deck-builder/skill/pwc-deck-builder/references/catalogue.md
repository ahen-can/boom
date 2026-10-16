# Pattern catalogue

Generated from assets/patterns.pptx by `inventory.py --catalogue`. Do not edit by hand: change the
pattern slide's speaker notes (PATTERN / PURPOSE / USE WHEN / ITEMS) and regenerate.

Slots: fill these keys in plan.json. `~N words` is a rough capacity at the template font size.
Item slots (`itemN_...`) can be left out from the end: unused items are removed with their shapes.

## TITLE
- **Purpose:** Cover slide
- **Use When:** First slide of every deck
- **Slots:** `title` (~12 words), `subtitle` (~25 words), `date` (~11 words)

## SECTION
- **Purpose:** Section divider
- **Use When:** Start of each major section (3-6 per deck)
- **Slots:** `section_number` (~1 words), `section_title` (~10 words)

## OBJ_3COL
- **Purpose:** Three objectives or themes side by side, each with an icon
- **Use When:** Objectives, key themes, pillars when there are exactly 2-3 parallel ideas
- **Items:** 2-3
- **Slots:** `title` (~24 words), `tracker` (~18 words), `item1_head` (~9 words), `item1_body` (~80 words), `item2_head` (~9 words), `item2_body` (~80 words), `item3_head` (~9 words), `item3_body` (~80 words), `source` (~26 words)
- **Icons:** `item1_icon`, `item2_icon`, `item3_icon`

## OBJ_NUMBERED_4
- **Purpose:** Numbered list of objectives or priorities with a short explanation each
- **Use When:** Objectives or priorities when there are 3-4 items needing a sentence each
- **Items:** 3-4
- **Slots:** `title` (~24 words), `tracker` (~18 words), `item1_head` (~17 words), `item1_body` (~58 words), `item2_head` (~17 words), `item2_body` (~58 words), `item3_head` (~17 words), `item3_body` (~58 words), `item4_head` (~17 words), `item4_body` (~58 words), `source` (~26 words)

## APPROACH_CHEVRON_4
- **Purpose:** Phased approach, left to right, with activities under each phase
- **Use When:** Approach, methodology or roadmap with 3-4 sequential phases
- **Items:** 3-4
- **Slots:** `title` (~24 words), `tracker` (~18 words), `item1_phase` (~15 words), `item1_time` (~6 words), `item1_body` (~85 words), `item2_phase` (~15 words), `item2_time` (~6 words), `item2_body` (~85 words), `item3_phase` (~15 words), `item3_time` (~6 words), `item3_body` (~85 words), `item4_phase` (~15 words), `item4_time` (~6 words), `item4_body` (~85 words), `source` (~26 words)

## TWO_COL
- **Purpose:** Two contrasting or complementary blocks
- **Use When:** Current vs future state, challenge vs response, what we heard vs what it means
- **Slots:** `title` (~24 words), `tracker` (~18 words), `left_head` (~8 words), `left_body` (~189 words), `right_head` (~8 words), `right_body` (~189 words), `source` (~26 words)

