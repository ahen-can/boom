# plan.json format

`build_deck.py` reads one JSON file that describes the whole deck.

```json
{
  "output": "Client_topic_draft.pptx",
  "defaults": {"source": ""},
  "slides": [
    {
      "pattern": "OBJ_3COL",
      "section": "Objectives",
      "fill": {
        "title": "Three outcomes will tell the client the programme worked",
        "item1_head": "Close in 6 days",
        "item1_body": "Cut the close from 12 to 6 working days by the March 2027 year end.",
        "item2_head": "One version of the numbers",
        "item2_body": ["Bullet one", "Bullet two"],
        "source": "Source: Client RFP (Sept 2026), section 3"
      },
      "icons": {"item1_icon": "target", "item2_icon": "data analytics"},
      "notes": "Speaker notes for the presenter."
    }
  ]
}
```

## Fields

| Field | Required | Meaning |
|---|---|---|
| `output` | no | File name for the deck. Saved to `/mnt/data/` in ChatPwC. |
| `defaults` | no | Slot values used on every slide unless the slide sets them (for example `source`, `client`). |
| `slides[].pattern` | yes | A pattern ID from `catalogue.md`. |
| `slides[].section` | no | Fills the `{{tracker}}` slot (the small breadcrumb above the title). |
| `slides[].fill` | yes | Slot name → text. A list of strings becomes one paragraph per item (bullets keep the template's formatting). For a table slot (`t1`), a list of rows, each a list of cell texts, header row first. |
| `slides[].icons` | no | Icon shape name → icon name or keywords. Matched against `assets/icons/icons_index.csv`. |
| `slides[].notes` | no | Speaker notes. |

## Rules

- Use only slot names listed for that pattern in `catalogue.md`. `build_deck.py --validate-only` lists any wrong names.
- Fewer items than the pattern holds: leave out the highest items (for example `item4_*`). Their shapes are removed.
- Any slot you leave out is removed if the shape holds nothing else, so do not pass empty strings to "clear" text.
- Keep within the `~N words` capacity in the catalogue. The builder shrinks text by at most 2pt and then reports overflow.
- Tables: give the header row and as many body rows as you need, up to the pattern's row count. Rows you
  leave out are removed. Use "" for an empty cell and "✓" for a tick if the original used ticks.
  Example: `"t1": [["Activity", "Weekly", "Monthly"], ["Review progress", "✓", ""]]`
- Match the length and kind of text in the catalogue's quoted examples. Never copy their content: it
  comes from another client's deck.
- Text goes in exactly as written: no markdown, no `**bold**`, no leading `-` or `•` (the template adds bullets).
