# -*- coding: utf-8 -*-
"""Rasterise a PDF into page-01.png, page-02.png ... for visual inspection.

Uses pypdfium2 from the bundled runtime (no Poppler needed).

Usage:
    python pdf2png.py input.pdf output_dir [dpi]
"""
import os
import sys

import pypdfium2 as pdfium

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main(src, outdir, dpi=150):
    os.makedirs(outdir, exist_ok=True)
    pdf = pdfium.PdfDocument(src)
    for i in range(len(pdf)):
        page = pdf[i]
        img = page.render(scale=dpi / 72).to_pil()
        img.save(os.path.join(outdir, "page-%02d.png" % (i + 1)))
    print("rendered %d pages at %d dpi -> %s" % (len(pdf), dpi, outdir))


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit("usage: pdf2png.py input.pdf output_dir [dpi]")
    main(sys.argv[1], sys.argv[2], int(sys.argv[3]) if len(sys.argv) > 3 else 150)
