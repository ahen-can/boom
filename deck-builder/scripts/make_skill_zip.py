"""Package the pwc-deck-builder skill as a zip for ChatPwC's "Upload skill".

    python deck-builder/scripts/make_skill_zip.py                       # dummy assets from the repo
    python deck-builder/scripts/make_skill_zip.py --private private      # your real files (see docs/LAPTOP_SETUP.md)
    python deck-builder/scripts/make_skill_zip.py --patterns my_patterns.pptx --icons my_icons/ \
        --storylines storylines.md --style-guide style_guide.md --out dist/

Real files override the repo's dummy ones in the zip only; the repo is never changed. The pattern
catalogue is regenerated from whichever patterns.pptx goes in.
"""
import argparse
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
SKILL_SRC = ROOT / "skill" / "pwc-deck-builder"
NAME = "pwc-deck-builder"
LIMIT_MB = 50


def stage(args) -> Path:
    tmp = Path(tempfile.mkdtemp(prefix="skillzip_")) / NAME
    shutil.copytree(SKILL_SRC, tmp, ignore=shutil.ignore_patterns("__pycache__", "*.pyc", ".DS_Store"))
    priv = Path(args.private) if args.private else None

    def pick(explicit, private_name):
        if explicit:
            return Path(explicit)
        if priv and (priv / private_name).exists():
            return priv / private_name
        return None

    patterns = pick(args.patterns, "patterns.pptx")
    icons = pick(args.icons, "icons")
    storylines = pick(args.storylines, "storylines.md")
    style = pick(args.style_guide, "style_guide.md")
    rules = pick(args.writing_rules, "writing_rules.md")
    used = {}
    if patterns:
        shutil.copy2(patterns, tmp / "assets" / "patterns.pptx")
        used["patterns"] = patterns
    if icons:
        shutil.rmtree(tmp / "assets" / "icons")
        shutil.copytree(icons, tmp / "assets" / "icons")
        used["icons"] = icons
    for src, rel in ((storylines, "storylines.md"), (style, "style_guide.md"), (rules, "writing_rules.md")):
        if src:
            shutil.copy2(src, tmp / "references" / rel)
            used[rel] = src
    for k, v in used.items():
        print(f"  using your {k}: {v}")
    if not used:
        print("  using the repo's dummy assets")
    return tmp


def validate(skill: Path):
    sys.path.insert(0, str(skill / "scripts"))
    import deckkit as dk  # noqa: E402

    prs = dk.open_deck(skill / "assets" / "patterns.pptx")
    pats = dk.index_patterns(prs)  # raises on duplicate IDs
    if not pats:
        raise SystemExit("patterns.pptx has no slides with 'PATTERN: <ID>' in the notes")
    missing = [i + 1 for i, s in enumerate(prs.slides) if not dk.pattern_id(s)]
    if missing:
        print(f"  warning: slides without a PATTERN note are ignored: {missing}")
    icons_dir = skill / "assets" / "icons"
    if icons_dir.is_dir():
        idx = dk.load_icon_index(icons_dir)
        absent = [r["file"] for r in idx if not (icons_dir / r["file"]).is_file()]
        if absent:
            print(f"  warning: icons_index.csv lists missing files: {absent[:10]}")
        print(f"  {len(idx)} icons")
    subprocess.run([sys.executable, str(skill / "scripts" / "inventory.py"), str(skill / "assets" / "patterns.pptx"),
                    "--catalogue", "--write", str(skill / "references" / "catalogue.md")], check=True,
                   capture_output=True)
    print(f"  {len(pats)} patterns: {', '.join(sorted(pats))}")


def zip_dir(skill: Path, dest: Path):
    dest.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(skill.rglob("*")):
            if p.is_file() and "__pycache__" not in p.parts:
                z.write(p, Path(NAME) / p.relative_to(skill))
    mb = dest.stat().st_size / 1e6
    print(f"wrote {dest} ({mb:.1f} MB)")
    if mb > LIMIT_MB:
        raise SystemExit(f"zip is over ChatPwC's {LIMIT_MB} MB limit: compress images in patterns.pptx "
                         "(File > Compress Pictures) or trim the icon set")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--private", help="folder with patterns.pptx, icons/, storylines.md, style_guide.md")
    ap.add_argument("--patterns")
    ap.add_argument("--icons")
    ap.add_argument("--storylines")
    ap.add_argument("--style-guide")
    ap.add_argument("--writing-rules")
    ap.add_argument("--out", default=str(ROOT / "dist"))
    a = ap.parse_args()
    skill = stage(a)
    validate(skill)
    zip_dir(skill, Path(a.out) / f"{NAME}.zip")
    shutil.rmtree(skill.parent, ignore_errors=True)


if __name__ == "__main__":
    main()
