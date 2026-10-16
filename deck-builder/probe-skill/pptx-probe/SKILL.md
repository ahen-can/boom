---
name: pptx-probe
description: Diagnostic skill. Use only when the user asks to run the pptx probe. Runs a Python script that reports what the code environment can do for PowerPoint generation (bundled skill files, python-pptx, LibreOffice, network, file output).
---

# pptx probe

Diagnostic only. Do not build a presentation or give advice. Run the probe and report.

## Steps

1. Locate this skill's bundled files in the code environment. Look for `scripts/probe.py` and
   `assets/marker.txt` (search the filesystem if needed, e.g. `find / -name marker.txt -path "*assets*" 2>/dev/null`).
2. If you found `scripts/probe.py`, run it with Python exactly as it is.
   If you could not find it, run the inline copy below instead, unchanged, and say clearly that the
   bundled files were NOT found.
3. Show the user the complete JSON output in a code block. Do not summarise or trim it.
4. Offer the files it wrote (`probe_xml.pptx`, `probe_pptx.pptx`) for download. They are in the
   folder given by `output_folder` in the report.
5. Below the JSON, add one line per item in `summary`: the key and yes/no. Nothing else.

## Inline copy of scripts/probe.py

```python
"""Probe what a code sandbox can do for PowerPoint generation.

Standard library only. Prints one JSON report and writes up to two test decks:
  probe_xml.pptx   - made by editing the slide XML directly (no python-pptx needed)
  probe_pptx.pptx  - made with python-pptx, if it is installed or installable

Usage: python probe.py [--no-install]
"""
import importlib
import json
import os
import platform
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.request
import zipfile
from pathlib import Path

MARKER_NAME = "marker.txt"
MARKER_TEXT = "pptx-probe-marker-7f3a"
TOKEN = "{{probe_token}}"
SEARCH_ROOTS = ["/mnt/data", "/mnt", "/skills", "/workspace", "/app", "/home", "/tmp", os.getcwd()]
LIBS = ["pptx", "lxml", "PIL", "docx", "openpyxl", "pandas", "matplotlib", "yaml"]
TOOLS = ["soffice", "libreoffice", "pdftoppm", "node", "pip"]

report = {"started": time.strftime("%Y-%m-%d %H:%M:%S")}


def step(name):
    """Run a check and store its result or its error, so one failure never stops the probe."""
    def wrap(fn):
        try:
            report[name] = fn()
        except Exception as e:  # noqa: BLE001
            report[name] = {"error": f"{type(e).__name__}: {e}"}
        return fn
    return wrap


def is_writable(folder):
    try:
        with tempfile.NamedTemporaryFile(dir=folder):
            return True
    except Exception:  # noqa: BLE001
        return False


def find_skill_root():
    """The folder holding assets/marker.txt: next to this script first, then a shallow search."""
    try:
        here = Path(__file__).resolve().parent.parent
        if (here / "assets" / MARKER_NAME).is_file():
            return here, "next to probe.py"
    except NameError:  # running as pasted inline code, no __file__
        pass
    for root in SEARCH_ROOTS:
        if not os.path.isdir(root):
            continue
        base_depth = root.rstrip("/").count("/")
        for dirpath, dirnames, filenames in os.walk(root):
            if dirpath.count("/") - base_depth >= 5:
                dirnames[:] = []
                continue
            if MARKER_NAME in filenames:
                p = Path(dirpath) / MARKER_NAME
                try:
                    if MARKER_TEXT in p.read_text(errors="ignore"):
                        return p.parent.parent, f"found by search under {root}"
                except OSError:
                    pass
    return None, "not found"


def output_dir():
    for d in ["/mnt/data", os.getcwd(), tempfile.gettempdir()]:
        if os.path.isdir(d) and is_writable(d):
            return Path(d)
    return Path(tempfile.gettempdir())


@step("environment")
def _env():
    return {
        "python": sys.version.split()[0],
        "executable": sys.executable,
        "platform": platform.platform(),
        "cwd": os.getcwd(),
        "user": os.environ.get("USER") or os.environ.get("USERNAME"),
        "writable": {d: is_writable(d) for d in ["/mnt/data", os.getcwd(), tempfile.gettempdir()] if os.path.isdir(d)},
        "mnt_data_listing": sorted(os.listdir("/mnt/data"))[:50] if os.path.isdir("/mnt/data") else None,
    }


SKILL_ROOT, how_found = find_skill_root()


@step("skill_files")
def _skill():
    if SKILL_ROOT is None:
        return {"found": False, "note": "Bundled skill files are not visible to code. Skills may deliver instructions only."}
    files = [str(p.relative_to(SKILL_ROOT)) for p in SKILL_ROOT.rglob("*") if p.is_file()]
    return {"found": True, "path": str(SKILL_ROOT), "how": how_found,
            "read_only": not is_writable(SKILL_ROOT), "files": sorted(files)[:50]}


def lib_versions():
    out = {}
    for name in LIBS:
        try:
            m = importlib.import_module(name)
            out[name] = getattr(m, "__version__", "installed")
        except Exception as e:  # noqa: BLE001
            out[name] = f"missing ({type(e).__name__})"
    return out


@step("libraries")
def _libs():
    return lib_versions()


@step("tools")
def _tools():
    return {t: shutil.which(t) for t in TOOLS}


@step("network")
def _net():
    try:
        with urllib.request.urlopen("https://pypi.org/simple/python-pptx/", timeout=5) as r:
            return {"pypi_reachable": True, "status": r.status}
    except Exception as e:  # noqa: BLE001
        return {"pypi_reachable": False, "error": f"{type(e).__name__}: {e}"[:200]}


@step("pip_install")
def _pip():
    if not str(report["libraries"].get("pptx", "")).startswith("missing"):
        return "skipped: python-pptx already installed"
    if "--no-install" in sys.argv:
        return "skipped: --no-install"
    if not report["network"].get("pypi_reachable"):
        return "skipped: no network to PyPI"
    r = subprocess.run([sys.executable, "-m", "pip", "install", "--quiet", "python-pptx"],
                       capture_output=True, text=True, timeout=90)
    importlib.invalidate_caches()
    return {"returncode": r.returncode, "stderr_tail": r.stderr[-300:],
            "pptx_after": lib_versions()["pptx"]}


OUT = output_dir()


def source_deck():
    """The bundled test deck, or a tiny one built from scratch if the skill files are missing."""
    if SKILL_ROOT and (SKILL_ROOT / "assets" / "probe.pptx").is_file():
        return SKILL_ROOT / "assets" / "probe.pptx", "bundled assets/probe.pptx"
    try:
        from pptx import Presentation
    except Exception:  # noqa: BLE001
        return None, "bundled deck not available and python-pptx missing"
    prs = Presentation()
    prs.slides.add_slide(prs.slide_layouts[0]).shapes.title.text = TOKEN
    path = OUT / "probe_source.pptx"
    prs.save(str(path))
    return path, "built from scratch (bundled deck not available)"


@step("xml_path")
def _xml():
    src, origin = source_deck()
    if src is None:
        return {"ok": False, "note": origin}
    dest = OUT / "probe_xml.pptx"
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zout:
        swapped = False
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "ppt/slides/slide1.xml" and TOKEN.encode() in data:
                data = data.replace(TOKEN.encode(), "XML edit worked".encode())
                swapped = True
            zout.writestr(item, data)
    return {"ok": swapped, "source": origin, "output": str(dest)}


@step("python_pptx_path")
def _pp():
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
    except Exception as e:  # noqa: BLE001
        return {"ok": False, "note": f"python-pptx not importable: {type(e).__name__}"}
    src, origin = source_deck()
    prs = Presentation(str(src)) if src else Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5] if len(prs.slide_layouts) > 5 else prs.slide_layouts[0])
    if slide.shapes.title is not None:
        slide.shapes.title.text = "python-pptx worked"
    box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(6), Inches(1))
    box.name = "probe_box"
    box.text_frame.text = f"Written {time.strftime('%Y-%m-%d %H:%M')}"
    box.text_frame.paragraphs[0].runs[0].font.size = Pt(18)
    dest = OUT / "probe_pptx.pptx"
    prs.save(str(dest))
    reread = Presentation(str(dest))
    names = [s.name for s in reread.slides[-1].shapes]
    return {"ok": "probe_box" in names, "source": origin, "output": str(dest),
            "slides": len(reread.slides), "shape_names_on_new_slide": names}


report["output_folder"] = str(OUT)
report["summary"] = {
    "skill_files_visible": bool(SKILL_ROOT),
    "python_pptx": not str(lib_versions().get("pptx", "missing")).startswith("missing"),
    "xml_edit_ok": isinstance(report.get("xml_path"), dict) and report["xml_path"].get("ok") is True,
    "pptx_write_ok": isinstance(report.get("python_pptx_path"), dict) and report["python_pptx_path"].get("ok") is True,
    "libreoffice": bool(report.get("tools", {}).get("soffice") or report.get("tools", {}).get("libreoffice")),
    "pypi_reachable": bool(report.get("network", {}).get("pypi_reachable")),
}
print(json.dumps(report, indent=2, default=str))
```
