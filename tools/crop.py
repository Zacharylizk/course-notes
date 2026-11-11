# -*- coding: utf-8 -*-
"""Crop a figure out of a rendered slide, without opening an image editor.

    python crop.py --src pages/page-07.png --box 0.06,0.42,0.55,0.93 --out img/fig-07a.png
    python crop.py ... --trim --pad 0.01          # trim white margins afterwards
    python crop.py ... --maxw 1100                # cap output width

Box is given as fractions of the image: x0,y0,x1,y1 (0 = left/top, 1 = right/bottom).
"""
import argparse
import os

from PIL import Image, ImageChops


def trim(im, pad, thresh=246):
    gray = im.convert("L").point(lambda v: 255 if v < thresh else 0)
    box = gray.getbbox()
    if not box:
        return im
    x0, y0, x1, y1 = box
    px = int(round(pad * im.width))
    py = int(round(pad * im.height))
    return im.crop((max(0, x0 - px), max(0, y0 - py),
                    min(im.width, x1 + px), min(im.height, y1 + py)))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--src", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--box", required=True, help="x0,y0,x1,y1 as fractions")
    ap.add_argument("--pad", type=float, default=0.0)
    ap.add_argument("--trim", action="store_true")
    ap.add_argument("--maxw", type=int, default=0)
    args = ap.parse_args()

    im = Image.open(args.src).convert("RGB")
    x0, y0, x1, y1 = [float(v) for v in args.box.split(",")]
    box = (int(x0 * im.width), int(y0 * im.height), int(x1 * im.width), int(y1 * im.height))
    im = im.crop(box)
    if args.trim:
        im = trim(im, args.pad)
    if args.maxw and im.width > args.maxw:
        im = im.resize((args.maxw, int(round(im.height * args.maxw / im.width))), Image.LANCZOS)
    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    im.save(args.out, optimize=True)
    print("%s  %dx%d" % (args.out, im.width, im.height))


if __name__ == "__main__":
    main()
