# Writing rules

Every headline, box and note in the deck follows these rules. Most come from the `unslop` skill in
pstack by Lauren Tan (MIT licence, github.com/backnotprop/pstack), adapted for slides.
`check_deck.py` fails the deck on any phrase in the banned list at the end.

## Headlines
- A headline states the point of the slide as a full sentence: "Month-end close takes 12 days, twice the peer median", not "Current challenges".
- One idea per headline, 15 words at most. Sentence case. No full stop at the end.
- Put the number in the headline when there is one.
- Read the headlines alone, in order. They should tell the whole story.

## Body text
- Box text can be a phrase, but it must read as normal English: keep articles and verbs. No arrows, no symbol-speak, no unexplained abbreviations.
- Name the mechanism or the number, not the feeling. "Cut the close from 12 to 6 days", not "a faster, smoother close".
- Use the client's own words and numbers from their material. Generic text that would fit any client gets cut.
- Plain words: use (not leverage or utilise), help (not facilitate), many (not numerous), if (not in the event that), to (not in order to), because (not due to the fact that).
- Active voice. Name who does what: "Finance owns the new process", not "The new process will be owned".
- Cut adverbs that prop up weak verbs. "Significantly improves" becomes the measured change.
- Use the natural number of items. Three columns only when there are three real things.
- Pick one term for one thing and repeat it. Do not cycle synonyms.
- No vague attributions ("experts believe", "industry reports suggest"). Name the source or cut it.
- No trailing "-ing" add-ons ("..., ensuring alignment", "..., driving value").
- No "not just X, but Y". State the point.
- No generic closers ("The future looks bright", "Together we can..."). End on a specific next step.

## Punctuation and format
- No em dashes. Use a full stop or a comma.
- Colons only before a list or an example.
- No emoji. Straight quotes. No bold inside body text unless the pattern already has it.
- Numbers: digits for 10 and above and for all measures (6 days, 3 FTEs). Currency and units as the client writes them.
- Every slide with figures has a source line: "Source: <document>, <section or page>; <date>".

## Banned phrases

```banned
leverage
leveraging
synergy
synergies
seamless
seamlessly
holistic
robust
cutting-edge
best-in-class
world-class
game-changer
game-changing
unlock
unlocking
unleash
empower
empowering
delve
tapestry
testament
pivotal
paradigm
north star
flywheel
in today's fast-paced
ever-evolving
rapidly evolving landscape
navigate the complexities
it is important to note
in order to
due to the fact that
at the end of the day
move the needle
drive value
value-add
best practices
going forward
```
