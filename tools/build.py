# -*- coding: utf-8 -*-
"""Assemble a lecture body fragment into the shared study-page shell.

Usage:
    python build.py body.html out.html [--course "5104MECH"] [--lecture "Lecture 3"]

The body fragment contains everything from the hero section onwards. This script:
  * derives the <title> from the first <h1>
  * builds the sidebar TOC from <h2 id="..."> headings
  * fills {{TITLE}}, {{NAV}}, {{COURSE}}, {{LECTURE}}, {{BODY}}
  * sanity-checks anchors, duplicate ids and local <img> paths
"""
import argparse
import html
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SHELL = os.path.join(os.path.dirname(HERE), "shell.html")
TAG_RE = re.compile(r"<[^>]+>")
WS_RE = re.compile(r"\s+")


def plain(fragment: str) -> str:
    txt = re.sub(r"<span[^>]*class=\"[^\"]*\ben\b[^\"]*\"[^>]*>.*?</span>", " ", fragment, flags=re.S)
    txt = TAG_RE.sub(" ", txt)
    return WS_RE.sub(" ", html.unescape(txt)).strip()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("body")
    ap.add_argument("out")
    ap.add_argument("--course", default="学习笔记 · Study Notes")
    ap.add_argument("--lecture", default="")
    args = ap.parse_args()

    body = open(args.body, encoding="utf-8").read()
    shell = open(SHELL, encoding="utf-8").read()

    m = re.search(r"<h1[^>]*>(.*?)</h1>", body, re.S)
    if not m:
        sys.exit("body has no <h1>")
    title = plain(m.group(1))

    nav, seen = [], set()
    for i, hm in enumerate(re.finditer(r"<h2([^>]*)>(.*?)</h2>", body, re.S), 1):
        attrs, inner = hm.group(1), hm.group(2)
        sec_start = body.rfind("<section", 0, hm.start())
        sec_tag = body[sec_start:body.find(">", sec_start) + 1] if sec_start != -1 else ""
        idm = re.search(r'id="([^"]+)"', attrs) or re.search(r'id="([^"]+)"', sec_tag)
        if not idm:
            print("  ! h2 without id skipped:", plain(inner)[:50])
            continue
        hid = idm.group(1)
        if hid in seen:
            sys.exit("duplicate h2 id: " + hid)
        seen.add(hid)
        num = re.search(r'<span[^>]*class="[^"]*\bn\b[^"]*"[^>]*>(.*?)</span>', inner, re.S)
        num_txt = plain(num.group(1)) if num else "%02d" % i
        label = plain(inner.replace(num.group(0), " ")) if num else plain(inner)
        nav.append('<a href="#%s"><span class="n">%s</span><span>%s</span></a>'
                   % (hid, html.escape(num_txt), html.escape(label)))

    out = (shell.replace("{{TITLE}}", html.escape(title))
                .replace("{{COURSE}}", html.escape(args.course))
                .replace("{{LECTURE}}", html.escape(args.lecture))
                .replace("{{NAV}}", "\n      ".join(nav))
                .replace("{{BODY}}", body))
    if "{{" in out:
        sys.exit("unfilled placeholder left in output")

    # --- checks ---------------------------------------------------------
    ids = re.findall(r'\bid="([^"]+)"', out)
    dupes = {x for x in ids if ids.count(x) > 1}
    if dupes:
        print("  ! duplicate ids:", ", ".join(sorted(dupes)))
    anchors = set(re.findall(r'href="#([^"]+)"', out))
    missing = anchors - set(ids)
    if missing:
        print("  ! nav targets missing:", ", ".join(sorted(missing)))
    outdir = os.path.dirname(os.path.abspath(args.out))
    for src in re.findall(r'<img[^>]+src="([^"]+)"', out):
        if src.startswith(("http://", "https://", "data:")):
            continue
        p = os.path.normpath(os.path.join(outdir, src))
        if not os.path.isfile(p):
            print("  ! missing image:", src)

    os.makedirs(outdir, exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as fh:
        fh.write(out)
    print("built %s  (%d sections, %.0f KB)" % (args.out, len(nav), len(out.encode("utf-8")) / 1024))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
