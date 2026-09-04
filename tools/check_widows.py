#!/usr/bin/env python3
"""Report section headings stranded at the foot of a column.

A title printed at the bottom of a column with its body over the break reads
badly, and it is invisible in the generator source and to every grep -- the only
evidence is glyph position on the rendered page. This reads word boxes out of the
PDF, clusters them into lines, splits each page into its columns, and reports any
heading that lands in the last few lines of a column.

  tools/check_widows.py documents/KC_Sourcebook.pdf [more.pdf ...]
  tools/check_widows.py            # every PDF in documents/

Two bands, on the same principle as check_columns.py -- a check that fails on
findings the reader will dismiss is a check the reader stops running:

  STRANDED   the heading is the last line in its column. Always wrong, and the
             build fails on it.
  tight      one or two lines of body under it. A heading wants at least three;
             two at a column break inside a page is ordinary and readable, so
             these are reported and not failed. Render and look.

Headings come from the matching corpus/*.md: ATX headings, plus stat-block names,
which are headings in every respect except their style.
"""
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
TIGHT = 2          # lines of body under a heading below which it is reported
LINE_TOL = 3.0     # points; words within this of each other share a line
FOOTER_GAP = 1.0   # points; the page-number footer sits below everything else


def headings_for(pdf):
    """Every string that acts as a heading in the document, from its corpus twin."""
    md = os.path.join(ROOT, "corpus", os.path.basename(pdf)[:-4] + ".md")
    if not os.path.exists(md):
        return set()
    out = set()
    for line in open(md, encoding="utf-8"):
        m = re.match(r'^#{1,6}\s+(.*?)\s*$', line)
        if not m:
            # A lone bold line is how SB() renders a stat-block name.
            m = re.match(r'^\*\*(.{1,60}?)\*\*\s*$', line)
        if m:
            h = re.sub(r'\*\*|\*|_', '', m.group(1)).strip()
            if h:
                out.add(h)
    return out


def pages_of(pdf):
    """[(page_width, [(xMin, xMax, yMax, text), ...]), ...] in reading order."""
    xml = subprocess.run(["pdftotext", "-bbox", pdf, "-"],
                         capture_output=True, text=True, check=True).stdout
    pages, cur = [], None
    for m in re.finditer(r'<page width="([\d.]+)"|'
                         r'<word xMin="([\d.]+)" yMin="[\d.]+" '
                         r'xMax="([\d.]+)" yMax="([\d.]+)">(.*?)</word>', xml):
        if m.group(1):
            cur = []
            pages.append((float(m.group(1)), cur))
        elif cur is not None:
            cur.append((float(m.group(2)), float(m.group(3)),
                        float(m.group(4)), unescape(m.group(5))))
    return pages


def unescape(s):
    for a, b in (("&lt;", "<"), ("&gt;", ">"), ("&quot;", '"'),
                 ("&apos;", "'"), ("&amp;", "&")):
        s = s.replace(a, b)
    return s


def columns(words, width):
    """The page's columns, left to right. A word straddling the middle means one."""
    mid = width / 2.0
    if any(w[0] < mid < w[1] for w in words):
        return [list(words)]
    left = [w for w in words if (w[0] + w[1]) / 2.0 < mid]
    right = [w for w in words if (w[0] + w[1]) / 2.0 >= mid]
    return [c for c in (left, right) if c]


def tail_lines(col, n):
    """The last n assembled lines of a column, top to bottom, footer excluded."""
    floor = max(w[2] for w in col)
    body = [w for w in col if w[2] < floor - FOOTER_GAP] or col
    ys = sorted({round(w[2], 1) for w in body})
    out = []
    for y in ys[-n:]:
        line = sorted((w for w in body if abs(w[2] - y) < LINE_TOL), key=lambda w: w[0])
        out.append(" ".join(w[3] for w in line).strip())
    return out


def main(argv):
    pdfs = argv or sorted(
        os.path.join(ROOT, "documents", f)
        for f in sorted(os.listdir(os.path.join(ROOT, "documents"))) if f.endswith(".pdf"))
    stranded, tight = [], []
    for pdf in pdfs:
        heads = headings_for(pdf)
        if not heads:
            continue
        name = os.path.basename(pdf)
        for n, (width, words) in enumerate(pages_of(pdf), 1):
            if not words:
                continue
            for ci, col in enumerate(columns(words, width)):
                lines = tail_lines(col, TIGHT + 1)
                for i, text in enumerate(lines):
                    if text not in heads:
                        continue
                    below = len(lines) - 1 - i
                    where = "%s p%-3d col%d  %-44r" % (name, n, ci, text)
                    (stranded if below == 0 else tight).append(
                        "%s  %d line(s) below" % (where, below))
    for line in tight:
        print("tight     " + line)
    for line in stranded:
        print("STRANDED  " + line)
    print("\n%d stranded, %d tight" % (len(stranded), len(tight)))
    return 1 if stranded else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
