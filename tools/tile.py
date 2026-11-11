# -*- coding: utf-8 -*-
"""Tile rendered pages into labelled contact sheets, so a 40-slide deck can be
scanned in six or seven reads instead of forty.

    python tile.py --dir work/lecture-03/pages --out work/lecture-03/sheets --scale 0.5
    python tile.py ... --pages 1-8 --per 4
"""
import argparse
import glob
import os
import re

from PIL import Image, ImageDraw, ImageFont

LABEL_H = 34
GAP = 10
FONT = r"C:\Windows\Fonts\arial.ttf"


def parse_pages(spec, available):
    if not spec:
        return available
    out = []
    for part in spec.split(","):
        part = part.strip()
        if "-" in part:
            a, b = part.split("-")
            out += list(range(int(a), int(b) + 1))
        elif part:
            out.append(int(part))
    return [p for p in out if p in set(available)]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dir", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--pages", default="")
    ap.add_argument("--per", type=int, default=4)
    ap.add_argument("--cols", type=int, default=2)
    ap.add_argument("--scale", type=float, default=0.5)
    args = ap.parse_args()

    files = glob.glob(os.path.join(args.dir, "page-*.png"))
    numbered = []
    for f in files:
        m = re.search(r"page-(\d+)\.png$", f)
        if m:
            numbered.append((int(m.group(1)), f))
    numbered.sort()
    by_num = dict(numbered)
    pages = parse_pages(args.pages, list(by_num))
    os.makedirs(args.out, exist_ok=True)
    font = ImageFont.truetype(FONT, 20) if os.path.exists(FONT) else None

    for start in range(0, len(pages), args.per):
        group = pages[start:start + args.per]
        ims = []
        for p in group:
            im = Image.open(by_num[p]).convert("RGB")
            if args.scale != 1:
                im = im.resize((max(1, int(im.width * args.scale)), max(1, int(im.height * args.scale))), Image.LANCZOS)
            ims.append((p, im))
        cw = max(im.width for _, im in ims)
        ch = max(im.height for _, im in ims)
        rows = (len(ims) + args.cols - 1) // args.cols
        sheet = Image.new("RGB", (args.cols * cw + (args.cols + 1) * GAP,
                                  rows * (ch + LABEL_H) + (rows + 1) * GAP), (255, 255, 255))
        draw = ImageDraw.Draw(sheet)
        for idx, (p, im) in enumerate(ims):
            r, c = divmod(idx, args.cols)
            x = GAP + c * (cw + GAP)
            y = GAP + r * (ch + LABEL_H + GAP)
            draw.rectangle([x, y, x + cw, y + LABEL_H - 5], fill=(32, 40, 52))
            draw.text((x + 9, y + 6), "p.%d" % p, fill=(255, 255, 255), font=font)
            sheet.paste(im, (x, y + LABEL_H))
        name = "sheet-p%02d-%02d.png" % (group[0], group[-1])
        path = os.path.join(args.out, name)
        sheet.save(path, optimize=True)
        print("%s  %dx%d  pages=%s" % (path, sheet.width, sheet.height,
                                       ",".join(str(g) for g in group)))


if __name__ == "__main__":
    main()
