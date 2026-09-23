#!/usr/bin/env python3
"""
check_page.py
=============

Confirms that a generated page is what it claims to be: one file that fetches
nothing, carries the workbook it was built from, and contains no em or en
dashes. Run it against the page before committing, or let the build workflow
run it for you.

    python3 tools/check_page.py index.html

Exits non-zero and names every problem it finds.
"""

import base64
import re
import sys

REMOTE = ("http://", "https://", "//", "ftp:", "ws:", "wss:")

# An asset attribute makes the browser fetch something in order to render the
# page. A hyperlink does not: it waits for a click. The first kind must never
# point off the file; the second may.
ASSET = re.compile(
    r'\b(src|srcset|poster|action|formaction|data|codebase|background|cite|manifest)'
    r'\s*=\s*["\']([^"\']+)["\']', re.I)
LINK = re.compile(r'<link\b[^>]*\bhref\s*=\s*["\']([^"\']+)["\']', re.I)
ANCHOR = re.compile(r'<a\b[^>]*\bhref\s*=\s*["\']([^"\']+)["\']', re.I)
DASHES = "‐‑‒–—―−"


def check(path):
    with open(path, encoding="utf-8") as fh:
        html = fh.read()
    problems = []

    for m in ASSET.finditer(html):
        if m.group(2).strip().startswith(REMOTE):
            problems.append("external %s reference: %s" % (m.group(1), m.group(2)[:90]))

    for m in LINK.finditer(html):
        if m.group(1).strip().startswith(REMOTE):
            problems.append("link tag pointing off the file: " + m.group(1)[:90])

    for tag in ("<iframe", "<object", "<embed", "<frame"):
        if tag in html.lower():
            problems.append("the page embeds %s, which can fetch from the network" % tag)

    for api in ("fetch(", "XMLHttpRequest", "importScripts", "WebSocket",
                "EventSource", "sendBeacon", "new Worker"):
        if api in html:
            problems.append("the page uses %s, which can reach the network" % api)

    present = [c for c in DASHES if c in html]
    if present:
        problems.append("dash characters present: %r" % present)

    nonascii = sorted({c for c in html if ord(c) > 127})
    if nonascii:
        problems.append("non-ascii characters present: %r" % nonascii)

    if "<title>" not in html:
        problems.append("the page has no title")
    if 'name="author" content="Alyssa I. Agard"' not in html:
        problems.append("the author meta tag is missing or has changed")
    if "creativecommons.org/licenses/by/4.0/" not in html:
        problems.append("the licence link is missing")
    if 'rel="icon"' not in html:
        problems.append("no favicon is embedded")

    m = re.search(r'const XLSX_B64 = "([A-Za-z0-9+/=]*)"', html)
    if not m:
        problems.append("the workbook is not carried inside the page")
    else:
        try:
            raw = base64.b64decode(m.group(1))
        except Exception as exc:
            problems.append("the embedded workbook does not decode: %s" % exc)
        else:
            if raw[:2] != b"PK":
                problems.append("the embedded workbook is not a zip container")
            elif len(raw) < 100000:
                problems.append("the embedded workbook is implausibly small: %d bytes" % len(raw))

    anchors = [m.group(1) for m in ANCHOR.finditer(html)
               if m.group(1).strip().startswith(REMOTE)]
    return problems, anchors, len(html.encode("utf-8"))


def main(argv):
    paths = argv[1:] or ["index.html"]
    failed = False
    for path in paths:
        problems, anchors, size = check(path)
        print("%s: %s bytes" % (path, format(size, ",")))
        for a in anchors:
            print("  hyperlink, followed only on a click: " + a)
        if problems:
            failed = True
            for p in problems:
                print("  PROBLEM: " + p)
        else:
            print("  self contained, no em or en dashes, workbook carried inside")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
