# Deck builder for ChatPwC

Goal: a ChatPwC skill + agent that turns RFPs and client material into PwC-style decks built from
human-made slide patterns. This folder currently holds Step 0, a probe that checks what ChatPwC's
code environment can do.

## Step 0: run the probe

1. Download `dist/pptx-probe.zip` (GitHub: open the file, then "Download raw file").
2. ChatPwC → Skills → **Upload skill** → `pptx-probe.zip`.
   If the upload is rejected, try `dist/pptx-probe-flat.zip` (same files, `SKILL.md` at the zip root).
3. Agent builder → new agent "Deck probe". Model: GPT-6.1 Sol. Turn on **Run Code** and **Skills**,
   then **Add Skills** → `pptx-probe`. Save.
4. In a chat with that agent: `Run the pptx probe and show me the full report.`
5. Copy the JSON report back to Claude, and download and open `probe_xml.pptx` and `probe_pptx.pptx`.

### Reading the summary

| Key | Meaning if `true` | If `false` |
|---|---|---|
| `skill_files_visible` | Scripts, template and icons can ship inside the skill zip | Ship assets via "Upload to Code Environment", or run locally |
| `python_pptx` | python-pptx is available (pre-installed or installed by the probe) | Builder uses the pure-XML path, or run locally |
| `xml_edit_ok` | Editing slide XML directly works | Something is wrong with file access |
| `pptx_write_ok` | Decks can be written with python-pptx | See `python_pptx` |
| `libreoffice` | Slides can be rendered to images for visual checks | You do the visual check in PowerPoint |
| `pypi_reachable` | Packages can be pip-installed | Only pre-installed packages can be used |

### Running locally instead

`python deck-builder/probe-skill/pptx-probe/scripts/probe.py` runs the same checks on your laptop
and writes the two test decks to the current folder.

## Rebuilding the zip

`python deck-builder/scripts/make_probe_zip.py` regenerates `assets/probe.pptx`, `SKILL.md`
(which embeds a copy of `probe.py`, so edit `probe.py` and rebuild) and both zips.
