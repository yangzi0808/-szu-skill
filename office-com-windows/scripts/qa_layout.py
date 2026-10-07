# -*- coding: utf-8 -*-
"""Layout QA for a rendered document: bottom whitespace + orphaned captions.

Run this on the PDF exported from the finished .docx. It catches the two
defects that are easy to ship by accident:

  * a large blank band at the bottom of a page because a figure/table did not
    fit and jumped to the next page;
  * a caption that landed on the next page while its figure stayed behind.

Usage:
    python qa_layout.py report.pdf [--gap-warn 25] [--caption-re "^[图表]"]
"""
import argparse
import sys

import numpy as np
import pdfplumber
import pypdfium2 as pdfium

sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def trailing_gaps(pdf_path, dpi=150):
    pdf = pdfium.PdfDocument(pdf_path)
    out = []
    for i in range(len(pdf)):
        img = pdf[i].render(scale=dpi / 72).to_pil().convert("L")
        a = np.asarray(img)
        h, w = a.shape
        # exclude the footer band (page number) and the outer borders
        body = a[: int(h * 0.918), int(w * 0.03): int(w * 0.97)]
        rows = np.where((body < 200).any(axis=1))[0]
        gap = 100.0 if len(rows) == 0 else (body.shape[0] - rows[-1]) / body.shape[0] * 100
        out.append(gap)
    return out


def orphan_captions(pdf_path):
    """A page whose topmost element is caption text, with an image starting lower."""
    problems = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            words = page.extract_words(use_text_flow=True)
            if not words:
                continue
            text_top = min(w["top"] for w in words)
            img_top = min((im["top"] for im in page.images), default=10 ** 9)
            if img_top <= text_top + 3:
                continue
            first = "".join(
                w["text"]
                for w in sorted(words, key=lambda w: w["x0"])
                if w["top"] <= text_top + 3
            )
            problems.append((i, first[:60]))
    return problems


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("pdf")
    ap.add_argument("--gap-warn", type=float, default=25.0,
                    help="report pages whose trailing whitespace exceeds this %%")
    ap.add_argument("--caption-re", default=r"^[图表]",
                    help="regex that identifies a caption line")
    args = ap.parse_args()

    import re
    capt = re.compile(args.caption_re)

    print("=== bottom whitespace per page ===")
    for i, gap in enumerate(trailing_gaps(args.pdf), 1):
        flag = "  <== large gap" if gap >= args.gap_warn else ""
        print("page %2d  trailing_gap=%5.1f%%%s" % (i, gap, flag))

    print("")
    print("=== orphaned captions ===")
    found = [p for p in orphan_captions(args.pdf) if capt.search(p[1])]
    for page_no, first in found:
        print("page %2d  <== starts with caption: %s" % (page_no, first))
    if not found:
        print("none")


if __name__ == "__main__":
    main()
