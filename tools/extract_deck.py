# -*- coding: utf-8 -*-
"""Turn a lecture deck (.pptx / .pdf) into an inspectable working folder.

    python extract_deck.py "D:\\path\\Lecture 3.pptx" --out "work/lecture-03"

Produces
    <out>/slides.json     structured text, image boxes, notes, per slide
    <out>/text.md         human/agent readable dump, one section per slide
    <out>/pages/page-NN.png   full-slide renders (PowerPoint COM / pdfium)
    <out>/images/…        embedded pictures extracted from a .pptx

Both a text layer and page renders are produced on purpose: lecture slides
routinely carry meaning in diagrams, arrows and equation images that text
extraction cannot see.  Read the renders when authoring the study page.
"""
import argparse
import glob
import hashlib
import html
import importlib
import json
import os
import re
import shutil
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
# 可选的备用解释器：缺库时本脚本会用它们重新跑自己（见 README「依赖」）。
# 默认指向 DSH 自带的运行时；装了别的 Python 就改这里，或用环境变量覆盖。
_HOME = os.environ.get("USERPROFILE") or os.path.expanduser("~")
DSH_PY = os.environ.get("CN_PY") or os.path.join(
    _HOME, ".dsh", "dsh-runtimes", "dsh-primary-runtime",
    "dependencies", "python", "python.exe")
PDF_PY = os.environ.get("CN_PY_ALT") or ""


def reexec(python_exe):
    if os.environ.get("_CN_REEXEC"):
        sys.exit("required library missing even after re-exec")
    os.environ["_CN_REEXEC"] = "1"
    cmd = [python_exe, os.path.abspath(__file__)] + sys.argv[1:]
    raise SystemExit(subprocess.call(cmd))


def need(module, fallback, why):
    try:
        return importlib.import_module(module)
    except ImportError:
        if fallback and os.path.isfile(fallback):
            print("  · %s missing here, re-running with %s" % (module, fallback))
            reexec(fallback)
        sys.exit("%s is required (%s)" % (module, why))


# --------------------------------------------------------------------------
# PPTX
# --------------------------------------------------------------------------
def pptx_pictures(shape, slide_no, img_dir, images):
    try:
        blob = shape.image.blob
        ext = shape.image.ext or "png"
    except Exception:
        return
    digest = hashlib.sha1(blob).hexdigest()[:8]
    name = "s%02d_%s.%s" % (slide_no, digest, ext.lower())
    path = os.path.join(img_dir, name)
    if not os.path.exists(path):
        with open(path, "wb") as fh:
            fh.write(blob)
    images.append({
        "file": "images/" + name,
        "left_cm": round(shape.left / 360000, 2) if shape.left is not None else None,
        "top_cm": round(shape.top / 360000, 2) if shape.top is not None else None,
        "width_cm": round(shape.width / 360000, 2) if shape.width is not None else None,
        "height_cm": round(shape.height / 360000, 2) if shape.height is not None else None,
    })


def pptx_shape(shape, slide_no, img_dir, images, blocks):
    from pptx.enum.shapes import MSO_SHAPE_TYPE

    if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
        for child in shape.shapes:
            pptx_shape(child, slide_no, img_dir, images, blocks)
        return

    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        for p in shape.text_frame.paragraphs:
            txt = (p.text or "").strip()
            if not txt:
                continue
            blocks.append({"kind": "text", "level": p.level, "text": txt})
        # OMML equations live outside the normal text runs
        xml = getattr(shape, "_element", None)
        if xml is not None:
            eqs = re.findall(r"<m:t[^>]*>(.*?)</m:t>", xml.xml, re.S)
            if eqs:
                eq = " ".join(html.unescape(re.sub(r"<[^>]+>", "", e)).strip() for e in eqs)
                eq = re.sub(r"\s+", " ", eq).strip()
                if eq:
                    blocks.append({"kind": "equation", "text": eq})

    if getattr(shape, "has_table", False) and shape.has_table:
        rows = []
        for row in shape.table.rows:
            cells = [re.sub(r"\s+", " ", c.text).strip() for c in row.cells]
            rows.append(" | ".join(cells))
        blocks.append({"kind": "table", "text": "\n".join(rows)})

    if getattr(shape, "has_chart", False) and shape.has_chart:
        ch = shape.chart
        title = ""
        try:
            if ch.has_title:
                title = ch.chart_title.text_frame.text
        except Exception:
            pass
        parts = ["[chart]" + (" " + title if title else "")]
        try:
            plot = ch.plots[0]
            cats = [str(c) for c in plot.categories]
            for s in plot.series:
                vals = [("" if v is None else v) for v in s.values]
                pairs = ", ".join("%s=%s" % (c, v) for c, v in zip(cats, vals))
                parts.append("%s: %s" % (s.name, pairs))
        except Exception:
            pass
        blocks.append({"kind": "chart", "text": "\n".join(parts)})

    try:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE or hasattr(shape, "image"):
            pptx_pictures(shape, slide_no, img_dir, images)
    except Exception:
        pass


def export_pptx_png(src, out_dir, width=1600, height=900):
    """Render slides with PowerPoint (late-bound VBScript, avoids broken PIAs)."""
    if os.name != "nt":
        return False
    pages = os.path.join(out_dir, "pages")
    os.makedirs(pages, exist_ok=True)
    vbs = os.path.join(out_dir, "_render.vbs")
    script = "\n".join([
        'On Error Resume Next',
        'Set ppt = CreateObject("PowerPoint.Application")',
        'If Err.Number <> 0 Then WScript.Echo "ERR|create|" & Err.Description : WScript.Quit 1',
        'ppt.Visible = True',
        'Set pres = ppt.Presentations.Open("%s", True, False, True)' % src,
        'If Err.Number <> 0 Then WScript.Echo "ERR|open|" & Err.Description : WScript.Quit 1',
        'pres.Export "%s", "PNG", %d, %d' % (pages, width, height),
        'If Err.Number <> 0 Then WScript.Echo "ERR|export|" & Err.Description',
        'pres.Close',
        'ppt.Quit',
        'WScript.Echo "OK"',
        '',
    ])
    with open(vbs, "w", encoding="utf-16") as fh:
        fh.write(script)
    try:
        r = subprocess.run(["cscript.exe", "//nologo", vbs], capture_output=True,
                           text=True, timeout=300)
        msg = (r.stdout or "").strip()
        if msg and not msg.endswith("OK"):
            print("  ! PowerPoint:", msg.replace("\n", " ")[:200])
    except Exception as exc:
        print("  ! PowerPoint render failed:", exc)
        return False
    finally:
        try:
            os.remove(vbs)
        except OSError:
            pass

    made = glob.glob(os.path.join(pages, "*.[pP][nN][gG]"))
    numbered = []
    for f in made:
        m = re.search(r"(\d+)\.png$", f, re.I)
        if m:
            numbered.append((int(m.group(1)), f))
    numbered.sort()
    if not numbered:
        return False
    for i, (_, f) in enumerate(numbered, 1):
        os.replace(f, os.path.join(pages, "page-%02d.png" % i))
    return True


def do_pptx(src, out_dir, render=True):
    pptx = need("pptx", DSH_PY, "python-pptx")
    from pptx import Presentation

    prs = Presentation(src)
    img_dir = os.path.join(out_dir, "images")
    os.makedirs(img_dir, exist_ok=True)
    slides = []
    for n, slide in enumerate(prs.slides, 1):
        blocks, images = [], []
        for shape in slide.shapes:
            try:
                pptx_shape(shape, n, img_dir, images, blocks)
            except Exception as exc:
                blocks.append({"kind": "error", "text": repr(exc)})
        title = ""
        try:
            if slide.shapes.title is not None:
                title = re.sub(r"\s+", " ", slide.shapes.title.text).strip()
        except Exception:
            pass
        notes = ""
        try:
            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text.strip()
        except Exception:
            pass
        slides.append({"n": n, "title": title, "blocks": blocks, "images": images, "notes": notes})

    meta = {
        "source": os.path.abspath(src),
        "kind": "pptx",
        "slide_width_cm": round(prs.slide_width / 360000, 2),
        "slide_height_cm": round(prs.slide_height / 360000, 2),
        "slides": len(slides),
        "text_layer": True,
    }
    if render:
        w = 1600
        h = int(round(w * prs.slide_height / prs.slide_width))
        meta["rendered"] = export_pptx_png(os.path.abspath(src), out_dir, w, h)
    else:
        meta["rendered"] = False

    json.dump({"meta": meta, "slides": slides},
              open(os.path.join(out_dir, "slides.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    write_text_md(out_dir, meta, slides)
    print("  slides: %d   renders: %s" % (len(slides), "yes" if meta.get("rendered") else "no"))


# --------------------------------------------------------------------------
# PDF
# --------------------------------------------------------------------------
def do_pdf(src, out_dir, render=True):
    pdfium = need("pypdfium2", PDF_PY, "PDF rendering")
    pdf = pdfium.PdfDocument(src)
    pages_dir = os.path.join(out_dir, "pages")
    os.makedirs(pages_dir, exist_ok=True)
    slides, total_chars = [], 0
    for n, page in enumerate(pdf, 1):
        w, h = page.get_size()
        try:
            text = page.get_textpage().get_text_range()
        except Exception:
            text = ""
        total_chars += len(text.strip())
        if render:
            scale = 1600.0 / max(w, 1)
            img = page.render(scale=scale).to_pil()
            img.save(os.path.join(pages_dir, "page-%02d.png" % n), optimize=True)
        blocks = []
        for para in re.split(r"\n\s*\n", text):
            para = re.sub(r"[ \t]+", " ", para).strip()
            if para:
                blocks.append({"kind": "text", "text": para})
        slides.append({"n": n, "title": "", "blocks": blocks, "images": [], "notes": "",
                       "width_pt": round(w, 1), "height_pt": round(h, 1)})
    meta = {
        "source": os.path.abspath(src),
        "kind": "pdf",
        "slides": len(slides),
        "text_layer": total_chars > 40,
        "rendered": bool(render),
    }
    if total_chars <= 40:
        print("  ! this PDF has (almost) no text layer — read pages/*.png as images")
    json.dump({"meta": meta, "slides": slides},
              open(os.path.join(out_dir, "slides.json"), "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)
    write_text_md(out_dir, meta, slides)
    print("  pages: %d   text chars: %d   renders: %s"
          % (len(slides), total_chars, "yes" if render else "no"))


# --------------------------------------------------------------------------
# shared
# --------------------------------------------------------------------------
def write_text_md(out_dir, meta, slides):
    lines = ["# %s" % os.path.basename(meta["source"]), "",
             "- kind: **%s**" % meta["kind"],
             "- slides/pages: **%d**" % meta["slides"],
             "- text layer: **%s**" % ("yes" if meta.get("text_layer") else "no — use page renders"),
             ""]
    for s in slides:
        title = (" — " + s["title"]) if s.get("title") else ""
        lines.append("## Slide %d%s" % (s["n"], title))
        if s.get("images"):
            for im in s["images"]:
                lines.append("![fig](%s)  *(%.1f×%.1f cm)*" % (im["file"], im.get("width_cm") or 0,
                                                              im.get("height_cm") or 0))
        for b in s["blocks"]:
            kind = b["kind"]
            if kind == "text":
                lines.append(("  " * b.get("level", 0)) + "- " + b["text"].replace("\n", " / "))
            elif kind == "equation":
                lines.append("- **[EQ]** `%s`" % b["text"])
            elif kind == "table":
                lines.append("- **[TABLE]**")
                for row in b["text"].splitlines():
                    lines.append("  | " + row)
            else:
                lines.append("- **[%s]** %s" % (kind.upper(), b["text"].replace("\n", " / ")))
        if s.get("notes"):
            lines.append("")
            lines.append("> **Notes:** " + s["notes"].replace("\n", " "))
        lines.append("")
    open(os.path.join(out_dir, "text.md"), "w", encoding="utf-8").write("\n".join(lines))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("deck")
    ap.add_argument("--out", required=True)
    ap.add_argument("--no-render", action="store_true")
    args = ap.parse_args()
    src = os.path.abspath(args.deck)
    if not os.path.isfile(src):
        sys.exit("no such file: " + src)
    out_dir = os.path.abspath(args.out)
    for sub in ("pages", "images"):
        p = os.path.join(out_dir, sub)
        if os.path.isdir(p):
            shutil.rmtree(p)
    os.makedirs(out_dir, exist_ok=True)
    print("extracting:", os.path.basename(src))
    ext = os.path.splitext(src)[1].lower()
    if ext == ".pptx":
        do_pptx(src, out_dir, render=not args.no_render)
    elif ext == ".pdf":
        do_pdf(src, out_dir, render=not args.no_render)
    else:
        sys.exit("only .pptx and .pdf are supported")
    print("wrote", out_dir)


if __name__ == "__main__":
    main()
