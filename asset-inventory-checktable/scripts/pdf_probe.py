#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Probe / render Chinese-scanned PDFs (pypdfium2).

Usage:
    python pdf_probe.py meta   <dir-or-pdf> [<dir-or-pdf> ...]
    python pdf_probe.py render <pdf> <outdir> [scale=2.5]
    python pdf_probe.py text   <pdf> <out.txt>

meta    prints pages, text-layer size and SCANNED flag for every PDF found.
render  renders pages to PNG (scale 2.5 ≈ 180 DPI) for multimodal reading.
text    dumps the whole text layer to a UTF-8 file (read that file, not the console,
        which mangles Chinese output).
"""
import os
import sys

import pypdfium2 as pdfium


def _pdfs(target):
    if os.path.isdir(target):
        out = []
        for rel in sorted(os.listdir(target)):
            if rel.lower().endswith(".pdf"):
                out.append(os.path.join(target, rel))
        return out
    return [target]


def cmd_meta(targets):
    total = 0
    for t in targets:
        for p in _pdfs(t):
            doc = pdfium.PdfDocument(p)
            n = len(doc)
            chars = sum(len(doc[i].get_textpage().get_text_range()) for i in range(n))
            doc.close()
            total += 1
            print("FILE:", os.path.basename(p))
            print("  PAGES:", n, " SIZE_KB:", round(os.path.getsize(p) / 1024, 1),
                  " TEXT:", chars, " SCANNED:", chars < 20 * n)
    print("TOTAL PDF:", total)


def cmd_render(pdf, outdir, scale=2.5):
    os.makedirs(outdir, exist_ok=True)
    doc = pdfium.PdfDocument(pdf)
    base = os.path.splitext(os.path.basename(pdf))[0]
    for i in range(len(doc)):
        img = doc[i].render(scale=scale).to_pil()
        fn = os.path.join(outdir, base + "_p" + str(i + 1).zfill(2) + ".png")
        img.save(fn)
        print("SAVED", fn, img.size)
    doc.close()


def cmd_text(pdf, outfile):
    doc = pdfium.PdfDocument(pdf)
    parts = []
    for i in range(len(doc)):
        parts.append("----- page " + str(i + 1) + " -----")
        parts.append(doc[i].get_textpage().get_text_range())
    doc.close()
    with open(outfile, "w", encoding="utf-8") as fh:
        fh.write("\n".join(parts))
    print("WROTE", outfile, os.path.getsize(outfile))


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    cmd = sys.argv[1]
    if cmd == "meta":
        cmd_meta(sys.argv[2:])
    elif cmd == "render":
        scale = float(sys.argv[4]) if len(sys.argv) > 4 else 2.5
        cmd_render(sys.argv[2], sys.argv[3], scale)
    elif cmd == "text":
        cmd_text(sys.argv[2], sys.argv[3])
    else:
        print(__doc__)
        sys.exit(1)


if __name__ == "__main__":
    main()
