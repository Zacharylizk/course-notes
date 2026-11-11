# -*- coding: utf-8 -*-
"""Pack a built study page into ONE self-contained HTML file.

    python inline.py lectures/<course>/<slug>/index.html lectures/<course>/<slug>/index-single.html

Every local <img src="img/xxx.png"> becomes a `data:` URI, so the result can be
moved / e-mailed / dropped on a phone as a single file and still show every
figure.  Nothing else about the page changes.

Why this exists: the built page references `img/…` with relative paths.  Users
move the .html on its own all the time and then report "the pictures are gone".
The single-file build removes that failure mode.
"""
import argparse
import base64
import mimetypes
import os
import re
import sys

SRC_RE = re.compile(r'(?P<attr>\bsrc=")(?P<path>[^"]+)(?P<close>")')
EXTRA_MIME = {".svg": "image/svg+xml", ".webp": "image/webp", ".woff2": "font/woff2"}


def is_local(path):
    return not path.startswith(("http://", "https://", "data:", "//", "#", "mailto:"))


def resolve(path, base):
    for cand in (os.path.join(base, path), os.path.join(base, "img", path)):
        if os.path.isfile(cand):
            return cand
    return None


def mime_of(path):
    ext = os.path.splitext(path)[1].lower()
    return EXTRA_MIME.get(ext) or mimetypes.guess_type(path)[0] or "application/octet-stream"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--base", default="", help="base dir for relative paths (default: dir of src)")
    args = ap.parse_args()

    src = os.path.abspath(args.src)
    base = os.path.abspath(args.base) if args.base else os.path.dirname(src)
    html = open(src, encoding="utf-8").read()

    used, missing = {}, []

    def embed(m):
        path = m.group("path").strip()
        if not is_local(path):
            return m.group(0)
        real = resolve(path, base)
        if not real:
            missing.append(path)
            return m.group(0)
        if real not in used:
            used[real] = "data:%s;base64,%s" % (
                mime_of(real),
                base64.b64encode(open(real, "rb").read()).decode("ascii"),
            )
        return m.group("attr") + used[real] + m.group("close")

    out = SRC_RE.sub(embed, html)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or ".", exist_ok=True)
    open(args.out, "w", encoding="utf-8").write(out)

    left = len(re.findall(r'\bsrc="(?!data:)', out))
    print("  inlined: %d files -> %s" % (len(used), os.path.basename(args.out)))
    print("  size: %.2f MB   remaining external src: %d" % (len(out.encode("utf-8")) / 1048576, left))
    if missing:
        print("  ! not found (left as-is): " + ", ".join(missing[:6]))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
