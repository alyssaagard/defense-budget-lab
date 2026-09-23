#!/usr/bin/env python3
"""
fy2627_budget_lab.py
====================

Reads FY2026_FY2027_Defense_Budget_Analytics.xlsx and writes one self-contained
HTML page that works offline from a file:// URL.

    python3 fy2627_budget_lab.py --xlsx <path> --out <path> [--open]

The workbook is the single source of truth. This script never recomputes a
figure the workbook already carries; it reads cached formula results and lays
them out. Every sheet name and every header label it depends on is verified at
run time. If a sheet is missing, or a header label has moved or been reworded,
the script stops and names the sheet and the cell rather than reading a column
by position and quietly producing a wrong page.

Standard library plus openpyxl. No other dependency, at build time or in the
emitted page: no content delivery networks, no web fonts, no charting library,
no framework. Charts are hand-written inline SVG.
"""

import argparse
import base64
import datetime
import json
import os
import re
import sys
import unicodedata

try:
    import openpyxl
except ImportError:
    sys.stderr.write("openpyxl is required: pip install openpyxl\n")
    raise SystemExit(2)

from openpyxl.utils import get_column_letter as CL

VERSION = "2.0"


class WorkbookMismatch(SystemExit):
    """Raised when the workbook does not look the way this script requires."""

    def __init__(self, msg):
        SystemExit.__init__(self, "\n  WORKBOOK MISMATCH\n  " + msg + "\n")


def _norm(v):
    """Compare header labels without being defeated by spacing or case."""
    if v is None:
        return ""
    s = unicodedata.normalize("NFKC", str(v))
    s = s.replace(" ", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


class Reader(object):
    """Every read goes through here, so no column is ever taken on trust."""

    def __init__(self, path):
        self.path = path
        try:
            self.wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
        except Exception as exc:
            raise WorkbookMismatch("could not open %s: %s" % (path, exc))
        self.checks = 0

    def precheck(self):
        """openpyxl cannot evaluate formulas, so the workbook has to carry the
        results of its own last recalculation. A workbook written by a script and
        never opened carries none, and every figure on the page would be blank.
        Catch that once, here, with an instruction rather than a puzzle."""
        sentinels = [("Category_Summary", 14, 2, "selected systems"),
                     ("FY2027_Categories", 14, 3, "FY2027 discretionary request"),
                     ("Pay_Module", 83, 2, "strength-weighted model cost per member"),
                     ("Year_Compare", 115, 15, "FY2026 total on the joined spine")]
        dead = []
        for name, row, col, what in sentinels:
            if name not in self.wb.sheetnames:
                continue
            v = self.wb[name].cell(row, col).value
            if not isinstance(v, (int, float)):
                dead.append("%s!%s%d (%s)" % (name, CL(col), row, what))
        if len(dead) >= 2:
            raise WorkbookMismatch(
                "the workbook carries no calculated results.\n"
                "  These cells should hold numbers and hold nothing: %s.\n"
                "  openpyxl reads cached results and cannot evaluate formulas, so "
                "a workbook that has never been recalculated has nothing to read. "
                "Open %s in Excel or LibreOffice, let it recalculate, save it, and "
                "run this again." % ("; ".join(dead), os.path.basename(self.path)))

    def sheet(self, name):
        self.checks += 1
        if name not in self.wb.sheetnames:
            raise WorkbookMismatch(
                "sheet %r is not in %s.\n  Sheets present: %s"
                % (name, os.path.basename(self.path), ", ".join(self.wb.sheetnames))
            )
        return self.wb[name]

    def label(self, ws, row, col, expect):
        """Assert the label at one cell, naming the cell when it disagrees."""
        self.checks += 1
        got = ws.cell(row, col).value
        if _norm(got) != _norm(expect):
            raise WorkbookMismatch(
                "%s!%s%d should read %r but reads %r.\n"
                "  A header has moved or been reworded. Fix the workbook or this "
                "script, but do not let the page read the column by position."
                % (ws.title, CL(col), row, expect, got)
            )
        return got

    def startswith(self, ws, row, col, prefix):
        self.checks += 1
        got = ws.cell(row, col).value
        if not _norm(got).startswith(_norm(prefix)):
            raise WorkbookMismatch(
                "%s!%s%d should begin %r but reads %r."
                % (ws.title, CL(col), row, prefix, got)
            )
        return got

    def cols(self, ws, row, wanted, search_to=None):
        """Map {key: expected label} to {key: column index} by reading row `row`."""
        self.checks += 1
        last = search_to or ws.max_column
        present = {}
        for c in range(1, last + 1):
            v = ws.cell(row, c).value
            if v is not None and str(v).strip() != "":
                present.setdefault(_norm(v), c)
        out = {}
        missing = []
        for key, lab in wanted.items():
            c = present.get(_norm(lab))
            if c is None:
                missing.append(lab)
            else:
                out[key] = c
                self.checks += 1
        if missing:
            raise WorkbookMismatch(
                "row %d of %s does not carry these headers: %s\n  Labels found on "
                "that row: %s"
                % (
                    row,
                    ws.title,
                    "; ".join(repr(m) for m in missing),
                    "; ".join(sorted(repr(k) for k in present.keys())),
                )
            )
        return out

    def need_number(self, ws, row, col, what):
        v = ws.cell(row, col).value
        if v is None:
            raise WorkbookMismatch(
                "%s!%s%d (%s) has no cached value.\n"
                "  openpyxl cannot evaluate formulas, so the workbook must carry "
                "results from its last recalculation. Open it in Excel or "
                "LibreOffice, let it recalculate, save, and run this again."
                % (ws.title, CL(col), row, what)
            )
        if not isinstance(v, (int, float)):
            raise WorkbookMismatch(
                "%s!%s%d (%s) should be a number but holds %r."
                % (ws.title, CL(col), row, what, v)
            )
        return float(v)


PENDING = ("pending", "n/a", "-", "not reported", "")


def num(v):
    """A blank stays a blank. CRS printing nothing is not CRS printing zero."""
    if v is None:
        return None
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s == "" or s.lower() in PENDING:
        return None
    s = s.replace("$", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def txt(v):
    if v is None:
        return ""
    if isinstance(v, datetime.datetime):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, datetime.date):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()


def r2(v, nd=10):
    return None if v is None else round(v, nd)

# ---------------------------------------------------------------------------
# Systems_Master: the 97 FY2026 lines
# ---------------------------------------------------------------------------

FY26_COLS = {
    "id": "ID",
    "table": "CRS table",
    "cat": "Category",
    "sub": "Subcategory",
    "name": "System",
    "svc": "Service (analyst)",
    "product": "CRS product",
    "recon": "Recon. request ($B)",
    "reconQ": "Recon. qty",
    "base": "Base request ($B)",
    "baseQ": "Base qty",
    "hasc": "H.R. 3838 HASC ($B)",
    "hascQ": "HASC qty",
    "sasc": "S. 2296 SASC ($B)",
    "sascQ": "SASC qty",
    "ndaa": "Enacted NDAA P.L. 119-60 ($B)",
    "ndaaQ": "NDAA qty",
    "hac": "H.R. 4016 HAC ($B)",
    "sac": "S. 2572 SAC ($B)",
    "sacQ": "SAC qty",
    "app": "Enacted approp. P.L. 119-75 ($B)",
    "spend": "Spend plan P.L. 119-21 ($B)",
    "spendQ": "Spend plan qty",
    "a1": "A-1 recon. request ($B)",
    "a1Q": "A-1 recon. qty",
    "total": "Total request ($B)",
    "mandShare": "Recon. share of request",
    "ndaaDelta": "NDAA minus base ($B)",
    "appDelta": "Approp. minus base ($B)",
    "enactTotal": "Enacted total FY2026 ($B)",
    "mandDep": "Mandatory dependence",
    "authAction": "Authorizer action",
    "appAction": "Appropriator action",
    "unit": "Budget-line cost per unit, base ($M)",
    "shareSel": "Share of selected base request",
    "shareCat": "Share of category approp.",
    "desc": "Description",
    "notes": "Notes (CRS footnotes and observations, paraphrased)",
}

FY26_FIRST = 6
FY26_LAST = 102
FY26_COUNT = 97


def read_fy26(rd):
    ws = rd.sheet("Systems_Master")
    rd.startswith(ws, 1, 1, "Systems Master Dataset")
    rd.label(ws, 4, 1, "Identification")
    rd.label(ws, 4, 8, "FY2026 DOD request")
    rd.label(ws, 4, 12, "Authorization")
    rd.label(ws, 4, 18, "Appropriation")
    rd.label(ws, 4, 26, "Computed metrics")
    c = rd.cols(ws, 5, FY26_COLS)
    rows = []
    for r in range(FY26_FIRST, FY26_LAST + 1):
        ident = txt(ws.cell(r, c["id"]).value)
        if not ident:
            raise WorkbookMismatch(
                "Systems_Master!A%d is empty. Rows %d to %d must all carry an ID."
                % (r, FY26_FIRST, FY26_LAST)
            )
        d = {"row": r}
        for k, col in c.items():
            v = ws.cell(r, col).value
            if k in ("id", "table", "cat", "sub", "name", "svc", "product",
                     "authAction", "appAction", "desc", "notes"):
                d[k] = txt(v)
            else:
                d[k] = r2(num(v))
        rows.append(d)
    if len(rows) != FY26_COUNT:
        raise WorkbookMismatch(
            "Systems_Master should carry %d lines in rows %d to %d, found %d."
            % (FY26_COUNT, FY26_FIRST, FY26_LAST, len(rows))
        )
    return rows


# ---------------------------------------------------------------------------
# FY2027_Systems: the 104 FY2027 lines
# ---------------------------------------------------------------------------

FY27_COLS = {
    "id": "ID",
    "table": "CRS table",
    "cat": "Category",
    "section": "Section (CRS grouping)",
    "name": "System (as printed)",
    "product": "CRS product",
    "note": "Note",
    "inCat": "In category tables",
    "mand": "Mandatory request ($B)",
    "mandQ": "Mand. qty",
    "disc": "Discretionary request ($B)",
    "discQ": "Disc. qty",
    "house": "H.R. 8800 House ($B)",
    "houseQ": "House qty",
    "sasc": "S. 4784 SASC ($B)",
    "sascQ": "SASC qty",
    "ndaa": "Enacted NDAA ($B)",
    "ndaaQ": "NDAA qty",
    "hac": "H.R. 9495 HAC ($B)",
    "hacQ": "HAC qty",
    "sac": "SAC ($B)",
    "sacQ": "SAC qty",
    "app": "Enacted approp. ($B)",
    "appQ": "Approp. qty",
    "total": "Total request ($B)",
    "mandShare": "Mandatory share of request",
    "houseDelta": "House minus disc. request ($B)",
    "housePct": "House vs. disc. request (%)",
    "sascDelta": "SASC minus disc. request ($B)",
    "sascPct": "SASC vs. disc. request (%)",
    "hacDelta": "HAC minus disc. request ($B)",
    "hacPct": "HAC vs. disc. request (%)",
    "houseAction": "House action",
    "sascAction": "SASC action",
    "hacAction": "HAC action",
    "enactAuth": "Enacted authorization",
    "enactApp": "Enacted appropriation",
    "unit": "Budget-line cost per unit, disc. ($M)",
    "shareSel": "Share of selected disc. request",
    "shareCat": "Share of category disc. request",
    "printedAs": "Printed as",
    "page": "CRS page",
    "footnote": "Footnote text",
    "figures": "Figures printed on this line",
}

FY27_FIRST = 6
FY27_LAST = 109
FY27_COUNT = 104

TEXT27 = ("id", "table", "cat", "section", "name", "product", "note", "inCat",
          "houseAction", "sascAction", "hacAction", "enactAuth", "enactApp",
          "printedAs", "footnote")


def read_fy27(rd):
    ws = rd.sheet("FY2027_Systems")
    rd.startswith(ws, 1, 1, "FY2027 Systems Dataset")
    rd.label(ws, 4, 1, "Identification")
    rd.label(ws, 4, 9, "FY2027 DOD request")
    rd.label(ws, 4, 13, "Authorization")
    rd.label(ws, 4, 19, "Appropriation")
    rd.label(ws, 4, 25, "Computed metrics")
    rd.label(ws, 4, 41, "Provenance")
    c = rd.cols(ws, 5, FY27_COLS)
    rows = []
    for r in range(FY27_FIRST, FY27_LAST + 1):
        ident = txt(ws.cell(r, c["id"]).value)
        if not ident:
            raise WorkbookMismatch(
                "FY2027_Systems!A%d is empty. Rows %d to %d must all carry an ID."
                % (r, FY27_FIRST, FY27_LAST)
            )
        d = {"row": r}
        for k, col in c.items():
            v = ws.cell(r, col).value
            d[k] = txt(v) if k in TEXT27 else r2(num(v))
        rows.append(d)
    if len(rows) != FY27_COUNT:
        raise WorkbookMismatch(
            "FY2027_Systems should carry %d lines in rows %d to %d, found %d."
            % (FY27_COUNT, FY27_FIRST, FY27_LAST, len(rows))
        )
    # R49023 reports a request and three committee marks. Nothing has been
    # enacted, and the Senate appropriators had not reported. Those four columns
    # must be empty, or the page would show a settled figure that does not exist.
    for key, lab in (("ndaa", "Enacted NDAA ($B)"),
                     ("app", "Enacted approp. ($B)"),
                     ("sac", "SAC ($B)"),
                     ("sacQ", "SAC qty")):
        filled = [d["id"] for d in rows if d.get(key) is not None]
        if filled:
            raise WorkbookMismatch(
                "FY2027_Systems column %r carries figures on %d lines (%s). "
                "Nothing was enacted for FY2027 and the Senate appropriators had "
                "not reported, so this column is expected to be empty. If that has "
                "changed, update this script rather than the check."
                % (lab, len(filled), ", ".join(filled[:6]))
            )
    return rows

# ---------------------------------------------------------------------------
# Category_Summary and FY2027_Categories
# ---------------------------------------------------------------------------

CATS_ORDER = ["Aircraft", "C4I", "Space", "Ground", "Hypersonic",
              "Missile Defense", "Missiles & Munitions", "Shipbuilding & Maritime"]

CAT26_COLS = {
    "cat": "Category",
    "n": "Systems",
    "recon": "Recon. request (Tables 1-7)",
    "base": "Base request",
    "total": "Total request",
    "hasc": "H.R. 3838 (HASC)",
    "sasc": "S. 2296 (SASC)",
    "ndaa": "Enacted NDAA",
    "hac": "H.R. 4016 (HAC)",
    "sac": "S. 2572 (SAC)",
    "app": "Enacted appropriation",
    "spend": "Spend plan (P.L. 119-21)",
    "enactTotal": "Enacted total FY2026",
    "ndaaDelta": "NDAA minus base",
    "appDelta": "Approp. minus base",
    "appNdaa": "Approp. minus NDAA",
    "a1": "Recon. request (Table A-1)",
    "mandDep": "Mandatory dependence",
    "program": "Program-wide DOD request",
    "coverage": "Selected coverage",
    "plus": "Plus-ups",
    "cuts": "Cuts",
    "asked": "As requested",
}

# The tie-outs this page is built on. If the workbook stops agreeing with them,
# something upstream changed and the page must not be published.
CAT26_TIEOUTS = [
    ("n", 97.0, 0.0001, "selected systems"),
    ("base", 114.07, 0.006, "base request"),
    ("ndaa", 124.86, 0.006, "enacted NDAA"),
    ("app", 127.60, 0.006, "enacted appropriation"),
    ("spend", 47.48, 0.006, "reconciliation spend plan"),
    ("enactTotal", 175.08, 0.006, "enacted total FY2026"),
    ("mandDep", 0.2711903, 0.000001, "mandatory dependence"),
]


def read_cats26(rd):
    ws = rd.sheet("Category_Summary")
    rd.startswith(ws, 1, 1, "Category Summary")
    rd.startswith(ws, 4, 1, "A. Selected systems by CRS category")
    c = rd.cols(ws, 5, CAT26_COLS)
    rows = []
    for i, name in enumerate(CATS_ORDER):
        r = 6 + i
        rd.label(ws, r, c["cat"], name)
        rows.append({k: (txt(ws.cell(r, col).value) if k == "cat"
                         else r2(num(ws.cell(r, col).value)))
                     for k, col in c.items()})
    rd.label(ws, 14, c["cat"], "All selected systems")
    tot = {k: (txt(ws.cell(14, col).value) if k == "cat"
               else r2(num(ws.cell(14, col).value))) for k, col in c.items()}
    for key, want, tol, what in CAT26_TIEOUTS:
        got = tot.get(key)
        if got is None:
            rd.need_number(ws, 14, c[key], what)
        if abs(got - want) > tol:
            raise WorkbookMismatch(
                "Category_Summary!%s14 (%s) reads %r; this page is built on %s. "
                "The workbook and the page disagree, so nothing is published."
                % (CL(c[key]), what, got, want)
            )
    # Section B, by service
    rd.startswith(ws, 18, 1, "B. Selected systems by service")
    sc = rd.cols(ws, 19, {
        "svc": "Service", "n": "Systems", "base": "Base request",
        "recon": "Recon. request", "ndaa": "Enacted NDAA",
        "app": "Enacted appropriation", "spend": "Spend plan",
        "enactTotal": "Enacted total", "appDelta": "Approp. minus base",
        "mandDep": "Mandatory dependence", "share": "Share of enacted total"})
    svc = []
    r = 20
    while r <= ws.max_row:
        nm = txt(ws.cell(r, sc["svc"]).value)
        if not nm or _norm(nm) == "total":
            break
        svc.append({k: (txt(ws.cell(r, col).value) if k == "svc"
                        else r2(num(ws.cell(r, col).value))) for k, col in sc.items()})
        r += 1
    rd.label(ws, r, sc["svc"], "Total")
    # Section C, by subcategory
    rd.startswith(ws, 32, 1, "C. Selected systems by subcategory")
    uc = rd.cols(ws, 33, {
        "cat": "Category", "sub": "Subcategory", "n": "Systems",
        "base": "Base request", "ndaa": "Enacted NDAA",
        "app": "Enacted appropriation", "spend": "Spend plan",
        "enactTotal": "Enacted total", "appDelta": "Approp. minus base",
        "mandDep": "Mandatory dependence"})
    sub = []
    r = 34
    while r <= ws.max_row:
        nm = txt(ws.cell(r, uc["cat"]).value)
        if not nm:
            break
        sub.append({k: (txt(ws.cell(r, col).value) if k in ("cat", "sub")
                        else r2(num(ws.cell(r, col).value))) for k, col in uc.items()})
        r += 1
    return {"rows": rows, "total": tot, "services": svc, "subs": sub}


CAT27_COLS = {
    "cat": "Category",
    "n": "Lines",
    "disc": "Discretionary request, computed ($B)",
    "discN": "Discretionary request, CRS narrative ($B)",
    "discGap": "Gap ($B)",
    "mand": "Mandatory request, computed ($B)",
    "mandN": "Mandatory request, CRS narrative ($B)",
    "total": "Total request, computed ($B)",
    "totalN": "Total request, CRS narrative ($B)",
    "program": "Category-wide request, CRS narrative ($B)",
    "coverage": "Selected share of category-wide",
    "house": "House auth ($B)",
    "sasc": "SASC auth ($B)",
    "hac": "HAC approp ($B)",
    "houseDelta": "House minus request ($B)",
    "sascDelta": "SASC minus request ($B)",
    "hacDelta": "HAC minus request ($B)",
    "page": "CRS page",
}

CAT27_NARRATIVE = [59.9, 2.1, 18.1, 4.8, 4.9, 15.7, 36.1, 61.5]


def read_cats27(rd):
    ws = rd.sheet("FY2027_Categories")
    rd.startswith(ws, 1, 1, "FY2027 Category Summary")
    rd.startswith(ws, 4, 1, "A. The eight categories")
    c = rd.cols(ws, 5, CAT27_COLS)
    rows = []
    for i, name in enumerate(CATS_ORDER):
        r = 6 + i
        rd.label(ws, r, c["cat"], name)
        d = {k: (txt(ws.cell(r, col).value) if k == "cat"
                 else r2(num(ws.cell(r, col).value))) for k, col in c.items()}
        want = CAT27_NARRATIVE[i]
        if d["discN"] is None or abs(d["discN"] - want) > 0.0001:
            raise WorkbookMismatch(
                "FY2027_Categories!%s%d should carry the CRS narrative "
                "discretionary figure %s for %s but reads %r."
                % (CL(c["discN"]), r, want, name, d["discN"])
            )
        rows.append(d)
    rd.label(ws, 14, c["cat"], "All eight categories")
    tot = {k: (txt(ws.cell(14, col).value) if k == "cat"
               else r2(num(ws.cell(14, col).value))) for k, col in c.items()}
    if tot["n"] != 101:
        raise WorkbookMismatch(
            "FY2027_Categories!%s14 should count 101 lines inside the eight "
            "category tables but reads %r." % (CL(c["n"]), tot["n"]))
    # Appendix Table A-1
    rd.startswith(ws, 16, 1, "B. Appendix Table A-1")
    ac = rd.cols(ws, 20, {
        "name": "Initiative", "mand": "Mandatory request ($B)",
        "disc": "Discretionary request ($B)", "house": "House ($B)",
        "sasc": "SASC ($B)", "hac": "HAC ($B)", "enacted": "Enacted"})
    a1 = []
    r = 21
    while r <= ws.max_row:
        nm = txt(ws.cell(r, ac["name"]).value)
        if not nm or _norm(nm).startswith("total"):
            break
        a1.append({k: (txt(ws.cell(r, col).value) if k in ("name", "enacted")
                       else r2(num(ws.cell(r, col).value))) for k, col in ac.items()})
        r += 1
    rd.startswith(ws, r, ac["name"], "Total, Table A-1")
    a1tot = {k: (txt(ws.cell(r, col).value) if k in ("name", "enacted")
                 else r2(num(ws.cell(r, col).value))) for k, col in ac.items()}
    a1note = txt(ws.cell(17, 1).value)
    # Investment accounts
    rd.startswith(ws, 26, 1, "C. Investment accounts")
    ic = rd.cols(ws, 27, {"item": "Item", "amt": "Amount ($B)", "basis": "Basis"})
    inv = []
    r = 28
    while r <= ws.max_row:
        nm = txt(ws.cell(r, ic["item"]).value)
        if not nm:
            break
        inv.append({"item": nm,
                    "amt": r2(num(ws.cell(r, ic["amt"]).value)),
                    "basis": txt(ws.cell(r, ic["basis"]).value)})
        r += 1
    return {"rows": rows, "total": tot, "a1": a1, "a1total": a1tot,
            "a1note": a1note, "invest": inv}

# ---------------------------------------------------------------------------
# Year_Compare
# ---------------------------------------------------------------------------

CMP_COLS = {
    "id26": "FY2026 ID",
    "id27": "FY2027 ID",
    "cat": "Category",
    "name": "System (FY2027 label)",
    "alt": "FY2026 label, where different",
    "presence": "Present in",
    "basis": "Crosswalk basis",
    "base26": "FY2026 base request ($B)",
    "disc27": "FY2027 discretionary request ($B)",
    "dDisc": "Change ($B)",
    "dDiscPct": "Change (%)",
    "recon26": "FY2026 reconciliation request ($B)",
    "mand27": "FY2027 mandatory request ($B)",
    "dMand": "Change ($B)",
    "tot26": "FY2026 total request ($B)",
    "tot27": "FY2027 total request ($B)",
    "dTot": "Change ($B)",
    "dTotPct": "Change (%)",
    "ndaa26": "FY2026 enacted NDAA ($B)",
    "house27": "FY2027 House ($B)",
    "sasc27": "FY2027 SASC ($B)",
    "app26": "FY2026 enacted appropriation ($B)",
    "hac27": "FY2027 HAC ($B)",
    "note": "Note",
}

CMP_FIRST = 12
CMP_LAST = 114
CMP_DIST = {"Both years": 95, "FY2027 only": 6, "FY2026 only": 2}


def read_compare(rd):
    ws = rd.sheet("Year_Compare")
    rd.startswith(ws, 1, 1, "FY2026 and FY2027 Line by Line")
    rd.startswith(ws, 10, 1, "A. Programmes on the joined spine")
    notes = [txt(ws.cell(r, 1).value) for r in range(5, 10)
             if txt(ws.cell(r, 1).value)]
    # Three column headers repeat the words "Change ($B)" and "Change (%)", so
    # the label map cannot disambiguate them on its own. Anchor them by position
    # and then verify the label that sits there.
    c = {}
    fixed = [("id26", 1), ("id27", 2), ("cat", 3), ("name", 4), ("alt", 5),
             ("presence", 6), ("basis", 7), ("base26", 8), ("disc27", 9),
             ("dDisc", 10), ("dDiscPct", 11), ("recon26", 12), ("mand27", 13),
             ("dMand", 14), ("tot26", 15), ("tot27", 16), ("dTot", 17),
             ("dTotPct", 18), ("ndaa26", 19), ("house27", 20), ("sasc27", 21),
             ("app26", 22), ("hac27", 23), ("note", 24)]
    for key, col in fixed:
        rd.label(ws, 11, col, CMP_COLS[key])
        c[key] = col
    rows = []
    for r in range(CMP_FIRST, CMP_LAST + 1):
        pres = txt(ws.cell(r, c["presence"]).value)
        if not pres:
            raise WorkbookMismatch(
                "Year_Compare!%s%d is empty; every row from %d to %d must say "
                "which years the line appears in."
                % (CL(c["presence"]), r, CMP_FIRST, CMP_LAST))
        d = {}
        for k, col in c.items():
            v = ws.cell(r, col).value
            d[k] = (txt(v) if k in ("id26", "id27", "cat", "name", "alt",
                                    "presence", "basis", "note")
                    else r2(num(v)))
        rows.append(d)
    seen = {}
    for d in rows:
        seen[d["presence"]] = seen.get(d["presence"], 0) + 1
    if seen != CMP_DIST:
        raise WorkbookMismatch(
            "Year_Compare rows %d to %d should split %s but split %s."
            % (CMP_FIRST, CMP_LAST, CMP_DIST, seen))
    # The bridge from the two year totals to the change column
    rd.startswith(ws, 117, 1, "Reconciling the two year totals")
    rd.label(ws, 121, 1, "Step")
    rd.label(ws, 121, 2, "Amount ($B)")
    rd.label(ws, 121, 3, "Note")
    bridge = []
    for r in range(122, 131):
        step = txt(ws.cell(r, 1).value)
        if not step:
            continue
        bridge.append({"step": step,
                       "amt": r2(num(ws.cell(r, 2).value)),
                       "note": txt(ws.cell(r, 3).value)})
    rd.startswith(ws, 130, 1, "Difference")
    close = num(ws.cell(130, 2).value)
    if close is None or abs(close) > 0.0051:
        raise WorkbookMismatch(
            "Year_Compare!B130 is the bridge residual and should close to zero; "
            "it reads %r. The joined spine does not reconcile with the change "
            "column, so the page would show a comparison that does not add up."
            % close)
    # Section B: lines present in one year only
    hrow = None
    for r in range(130, min(ws.max_row, 155) + 1):
        if _norm(ws.cell(r, 1).value) == "direction":
            hrow = r
            break
    if hrow is None:
        raise WorkbookMismatch(
            "Year_Compare section B has no header row: no cell in column A "
            "between rows 130 and 155 reads 'Direction'.")
    oc = rd.cols(ws, hrow, {
        "dir": "Direction", "cat": "Category", "sys": "System",
        "tot26": "FY2026 total request ($B)", "tot27": "FY2027 total request ($B)"})
    only = []
    r = hrow + 1
    while r <= ws.max_row:
        dd = txt(ws.cell(r, oc["dir"]).value)
        if not dd:
            break
        only.append({"dir": dd,
                     "cat": txt(ws.cell(r, oc["cat"]).value),
                     "sys": txt(ws.cell(r, oc["sys"]).value),
                     "tot26": r2(num(ws.cell(r, oc["tot26"]).value)),
                     "tot27": r2(num(ws.cell(r, oc["tot27"]).value))})
        r += 1
    if len(only) != 8:
        raise WorkbookMismatch(
            "Year_Compare section B should list 8 lines that appear in one year "
            "only (6 new, 2 dropped); it lists %d." % len(only))
    return {"notes": notes, "rows": rows, "bridge": bridge, "only": only}

# ---------------------------------------------------------------------------
# The pay module: assumptions, grade rows, basic pay table, force, allowances
# ---------------------------------------------------------------------------

PAY_ASSUMPTIONS = [
    (16, "Fiscal year being priced"),
    (17, "Retired pay accrual rate"),
    (18, "Include retired pay accrual?"),
    (19, "TSP automatic agency contribution"),
    (20, "TSP matching contribution"),
    (21, "Total TSP agency contribution"),
    (22, "Employer Medicare (HI) rate"),
    (23, "Employer Social Security (OASDI) rate"),
    (24, "OASDI taxable maximum"),
    (25, "Monthly cap, O-6 and below"),
    (26, "Monthly cap, O-7 to O-10"),
    (27, "Raise effective month (1 = January)"),
    (28, "Months of the fiscal year affected"),
    (29, "Across-the-board raise applied to every grade"),
    (30, "Service scope for strength"),
    (31, "Service for the composite rate comparison"),
    (32, "Composite rate edition in use"),
]

PAYROW_COLS = {
    "g": "Pay grade",
    "group": "Group",
    "strength": "Active-duty strength",
    "yos": "Representative years of service",
    "pay": "Monthly basic pay, Jan 2026 ($)",
    "cap": "Statutory monthly cap ($)",
    "payroll": "Current annual basic payroll ($M)",
    "model": "Model burdened cost per member, full year ($)",
    "composite": "Composite standard rate, selected service ($)",
    "spread": "Spread, model minus composite ($)",
    "spreadPct": "Spread (%)",
    "bah": "Monthly housing per member ($)",
    "bas": "Monthly subsistence per member ($)",
    "txad": "Federal income tax advantage, current ($/yr)",
    "rmc": "Regular military compensation, current ($/yr)",
    "basis": "RMC basis",
}

GRADE_FIRST = 37
GRADE_LAST = 60
GRADE_COUNT = 24


def read_pay(rd):
    ws = rd.sheet("Pay_Module")
    rd.startswith(ws, 5, 1, "A. Rates derived from the M-1")
    rd.label(ws, 6, 1, "Item")
    rd.label(ws, 6, 2, "FY2026")
    rd.label(ws, 6, 3, "FY2027")
    rd.label(ws, 6, 4, "Unit")
    rd.label(ws, 6, 5, "Source")
    m1 = []
    for r in range(7, 13):
        item = txt(ws.cell(r, 1).value)
        if not item:
            continue
        m1.append({"item": item,
                   "fy26": r2(num(ws.cell(r, 2).value)),
                   "fy27": r2(num(ws.cell(r, 3).value)),
                   "unit": txt(ws.cell(r, 4).value),
                   "source": txt(ws.cell(r, 5).value)})
    rd.startswith(ws, 7, 1, "Active-duty basic pay")
    rd.startswith(ws, 8, 1, "Retired pay accrual")
    rd.startswith(ws, 9, 1, "TSP matching contributions")
    rd.label(ws, 15, 1, "Assumption")
    rd.label(ws, 15, 2, "Value")
    rd.label(ws, 15, 3, "Unit")
    rd.label(ws, 15, 4, "Source or basis")
    assum = []
    for r, lab in PAY_ASSUMPTIONS:
        rd.label(ws, r, 1, lab)
        v = ws.cell(r, 2).value
        assum.append({"k": lab,
                      "v": r2(num(v)) if isinstance(v, (int, float)) else txt(v),
                      "num": isinstance(v, (int, float)),
                      "unit": txt(ws.cell(r, 3).value),
                      "basis": txt(ws.cell(r, 4).value)})
    A = dict((a["k"], a["v"]) for a in assum)
    rd.startswith(ws, 35, 1, "A raise moves basic pay")
    c = rd.cols(ws, 36, PAYROW_COLS)
    grades = []
    for r in range(GRADE_FIRST, GRADE_LAST + 1):
        g = txt(ws.cell(r, c["g"]).value)
        if not g:
            raise WorkbookMismatch(
                "Pay_Module!A%d is empty; rows %d to %d must all carry a pay grade."
                % (r, GRADE_FIRST, GRADE_LAST))
        d = {}
        for k, col in c.items():
            v = ws.cell(r, col).value
            d[k] = txt(v) if k in ("g", "group", "yos", "basis") else r2(num(v))
        grades.append(d)
    if len(grades) != GRADE_COUNT:
        raise WorkbookMismatch(
            "Pay_Module should carry %d pay grades in rows %d to %d, found %d."
            % (GRADE_COUNT, GRADE_FIRST, GRADE_LAST, len(grades)))
    # Strength-weighted cross-checks
    rd.label(ws, 81, 1, "Measure")
    rd.label(ws, 81, 2, "Value")
    rd.label(ws, 82, 1, "Grades with both figures present")
    rd.label(ws, 83, 1, "Strength-weighted model cost per member ($)")
    rd.label(ws, 84, 1, "Strength-weighted composite standard rate ($)")
    wgt = {"n": r2(num(ws.cell(82, 2).value)),
           "model": r2(num(ws.cell(83, 2).value)),
           "composite": r2(num(ws.cell(84, 2).value))}
    if wgt["n"] != GRADE_COUNT:
        raise WorkbookMismatch(
            "Pay_Module!B82 should count %d grades carrying both a model cost and "
            "a composite rate but reads %r." % (GRADE_COUNT, wgt["n"]))
    if wgt["model"] is None or wgt["model"] <= 0:
        raise WorkbookMismatch(
            "Pay_Module!B83 (strength-weighted model cost per member) reads %r. "
            "It is blank or zero, which means the workbook was saved without a "
            "usable recalculation." % wgt["model"])
    return {"m1": m1, "assumptions": assum, "A": A, "grades": grades,
            "weighted": wgt}


def read_pay_table(rd):
    ws = rd.sheet("Pay_Table_2026")
    rd.startswith(ws, 1, 1, "2026 Monthly Basic Pay Table")
    rd.label(ws, 4, 1, "Monthly basic pay ($)")
    rd.label(ws, 5, 1, "Pay grade")
    rd.label(ws, 5, 2, "Group")
    steps = []
    for c in range(3, ws.max_column + 1):
        lab = txt(ws.cell(5, c).value)
        if not lab:
            break
        steps.append({"label": lab, "col": c})
    if len(steps) != 22:
        raise WorkbookMismatch(
            "Pay_Table_2026 row 5 should carry 22 years-of-service steps from "
            "C5; it carries %d." % len(steps))
    if steps[0]["label"] != "2 or less":
        raise WorkbookMismatch(
            "Pay_Table_2026!C5 should read '2 or less' but reads %r."
            % steps[0]["label"])
    table = {}
    order = []
    r = 6
    while r <= ws.max_row:
        g = txt(ws.cell(r, 1).value)
        if not g:
            break
        order.append(g)
        table[g] = [r2(num(ws.cell(r, s["col"]).value)) for s in steps]
        r += 1
    return {"steps": [s["label"] for s in steps], "grades": order, "table": table}


def read_force(rd):
    ws = rd.sheet("Force_Structure")
    rd.startswith(ws, 1, 1, "Active-Duty Force Structure by Pay Grade")
    rd.startswith(ws, 4, 1, "A. Enlisted personnel")
    c = rd.cols(ws, 5, {
        "g": "Pay grade", "Army": "Army", "Navy": "Navy",
        "Marine Corps": "Marine Corps", "Air Force": "Air Force",
        "Space Force": "Space Force", "total": "Total (computed)",
        "pub": "Total (published)", "check": "Check"})
    services = ["Army", "Navy", "Marine Corps", "Air Force", "Space Force"]
    force = {}
    bysvc = {}
    order = []

    def block(first, endlabel):
        r = first
        while r <= ws.max_row:
            g = txt(ws.cell(r, c["g"]).value)
            if not g:
                break
            if _norm(g).startswith(_norm(endlabel)):
                return r
            order.append(g)
            force[g] = r2(num(ws.cell(r, c["total"]).value))
            bysvc[g] = dict((s, r2(num(ws.cell(r, c[s]).value))) for s in services)
            r += 1
        raise WorkbookMismatch(
            "Force_Structure has no %r row after row %d." % (endlabel, first))

    r_enl = block(6, "Total enlisted")
    rd.startswith(ws, r_enl, c["g"], "Total enlisted")
    # The officer block repeats the same headers further down.
    ohdr = None
    for r in range(r_enl + 1, ws.max_row + 1):
        if _norm(ws.cell(r, c["g"]).value) == "pay grade":
            ohdr = r
            break
    if ohdr is None:
        raise WorkbookMismatch(
            "Force_Structure has no second 'Pay grade' header for the officer "
            "block after row %d." % r_enl)
    r_off = block(ohdr + 1, "Total officers")
    rd.startswith(ws, r_off, c["g"], "Total officers")
    tot = sum(v for v in force.values() if v)
    if abs(tot - 1285084) > 0.5:
        raise WorkbookMismatch(
            "Force_Structure pay grades sum to %s; this page is built on an "
            "active-duty total of 1,285,084." % format(int(tot), ","))
    return {"order": order, "total": force, "bysvc": bysvc,
            "services": services, "sum": tot}


def read_composite(rd):
    ws = rd.sheet("Pay_Composite")
    rd.startswith(ws, 1, 1, "Military Composite Standard Pay")
    rd.startswith(ws, 5, 1, "A. FY2026 rates")
    rd.startswith(ws, 33, 1, "B. FY2027 rates")
    services = ["Army", "Navy", "Marine Corps", "Air Force", "Space Force"]

    def block(hrow):
        want = {"g": "Pay grade"}
        for s in services:
            want[s + "|basic"] = "%s: average basic pay ($)" % s
            want[s + "|comp"] = "%s: DoD composite standard pay rate ($)" % s
        c = rd.cols(ws, hrow, want)
        out = {}
        r = hrow + 1
        while r <= ws.max_row:
            g = txt(ws.cell(r, c["g"]).value)
            if not g:
                break
            out[g] = dict(
                (s, {"basic": r2(num(ws.cell(r, c[s + "|basic"]).value)),
                     "comp": r2(num(ws.cell(r, c[s + "|comp"]).value))})
                for s in services)
            r += 1
        return out

    return {"fy26": block(6), "fy27": block(34), "services": services,
            "note26": txt(ws.cell(5, 1).value), "note27": txt(ws.cell(33, 1).value)}


ALLOW_COLS = {
    "g": "Pay grade",
    "group": "Group",
    "strength": "Active-duty strength",
    "cashShare": "Share drawing cash housing, CY2023",
    "married": "Married share, CY2023",
    "pubBah": "Published housing, 2023 ($/yr)",
    "pubBas": "Published subsistence, 2023 ($/yr)",
    "bahScale": "Housing scale factor",
    "basScale": "Subsistence scale factor",
    "bahYr": "Housing per member drawing cash, priced year ($/yr)",
    "basYr": "Subsistence rate, priced year ($/yr)",
    "allowYr": "Allowances, priced year ($/yr)",
    "basicYr": "Basic pay at the representative years of service ($/yr)",
    "rateSingle": "Marginal rate, single",
    "rateMarried": "Marginal rate, married",
    "rateBlend": "Blended marginal rate",
    "txad": "Federal income tax advantage ($/yr)",
    "rmc": "Regular military compensation ($/yr)",
    "basicShare": "Basic pay as a share of RMC",
    "pubRmc": "Published RMC, 1 January 2023 ($/yr)",
    "pubShare": "Published basic pay share of RMC, 2023",
    "pubAllow": "Published allowances, 2023 ($/yr)",
    "pubTxad": "Published tax advantage, 2023 ($/yr)",
    "pubRate": "Implied marginal rate behind the published tax advantage",
}


def read_allow(rd):
    ws = rd.sheet("Pay_Allowances")
    rd.startswith(ws, 1, 1, "Allowances and Regular Military Compensation")
    rd.startswith(ws, 5, 1, "A. The M-1's allowance dollars")
    rd.label(ws, 6, 1, "Item")
    rd.label(ws, 6, 2, "FY2026 ($K)")
    rd.label(ws, 6, 3, "FY2027 ($K)")
    rd.label(ws, 6, 4, "Source")
    m1 = []
    for r in range(7, 13):
        item = txt(ws.cell(r, 1).value)
        if not item:
            continue
        m1.append({"item": item,
                   "fy26": r2(num(ws.cell(r, 2).value)),
                   "fy27": r2(num(ws.cell(r, 3).value)),
                   "source": txt(ws.cell(r, 4).value)})
    rd.startswith(ws, 34, 1, "D. Regular military compensation, per grade")
    c = rd.cols(ws, 35, ALLOW_COLS)
    rows = []
    r = 36
    while r <= ws.max_row:
        g = txt(ws.cell(r, c["g"]).value)
        if not g or _norm(g) == "total":
            break
        d = {}
        for k, col in c.items():
            v = ws.cell(r, col).value
            d[k] = txt(v) if k in ("g", "group") else r2(num(v))
        rows.append(d)
        r += 1
    if len(rows) != GRADE_COUNT:
        raise WorkbookMismatch(
            "Pay_Allowances section D should carry %d pay grades from row 36; "
            "it carries %d." % (GRADE_COUNT, len(rows)))
    rd.label(ws, 71, 1, "Measure")
    rd.label(ws, 72, 1, "Strength-weighted basic pay share of RMC")
    rd.label(ws, 73, 1, "Strength-weighted implied marginal rate behind the tax advantage")
    rd.startswith(ws, 62, 1, "E. Reconciliation and diagnostics")
    rd.label(ws, 63, 1, "Check")
    rd.label(ws, 63, 2, "Value")
    rd.label(ws, 63, 3, "Target")
    recon = []
    r = 64
    while r <= ws.max_row:
        k = txt(ws.cell(r, 1).value)
        if not k or _norm(k) == "measure":
            break
        recon.append({"check": k,
                      "value": r2(num(ws.cell(r, 2).value)),
                      "target": r2(num(ws.cell(r, 3).value))})
        r += 1
    wgt = {"sharePriced": r2(num(ws.cell(72, 3).value)),
           "sharePub": r2(num(ws.cell(72, 2).value)),
           "ratePriced": r2(num(ws.cell(73, 3).value)),
           "ratePub": r2(num(ws.cell(73, 2).value)),
           "note": txt(ws.cell(70, 1).value)}
    return {"m1": m1, "rows": rows, "weighted": wgt, "recon": recon}

# ---------------------------------------------------------------------------
# Context: toplines, timeline, glossary, sources, QA, projection
# ---------------------------------------------------------------------------

def read_simple(rd, sheet, hrow, wanted, start, textkeys=(), stop=None):
    ws = rd.sheet(sheet)
    c = rd.cols(ws, hrow, wanted)
    first = list(wanted.keys())[0]
    out = []
    r = start
    while r <= ws.max_row:
        k = ws.cell(r, c[first]).value
        lab = txt(k)
        if not lab:
            break
        if stop and _norm(lab).startswith(_norm(stop)):
            break
        d = {}
        for key, col in c.items():
            v = ws.cell(r, col).value
            d[key] = (txt(v) if (key in textkeys or key == first or
                                 isinstance(v, str) or v is None)
                      else r2(num(v)))
        out.append(d)
        r += 1
    return out


def read_context(rd):
    out = {}

    # Toplines
    ws = rd.sheet("Topline")
    rd.startswith(ws, 1, 1, "DOD Toplines")
    top = {}
    for tag, hdr, first, title in (("fy26", 5, 6, "A. FY2026 budget request toplines"),
                                   ("recon", 18, 19, "B. Reconciliation funding"),
                                   ("supp", 39, 40, "C. FY2026 supplemental request"),
                                   ("ctx", 52, 53, "D. The selected systems in context")):
        rd.startswith(ws, hdr - 1, 1, title)
        rd.label(ws, hdr, 1, "Item")
        rd.label(ws, hdr, 3, "Amount ($B)")
        rd.label(ws, hdr, 4, "Basis")
        rows = []
        r = first
        while r <= ws.max_row:
            item = txt(ws.cell(r, 1).value)
            if not item:
                break
            rows.append({"item": item,
                         "amt": r2(num(ws.cell(r, 3).value)),
                         "basis": txt(ws.cell(r, 4).value),
                         "src": txt(ws.cell(r, 5).value)})
            r += 1
        top[tag] = rows
    out["topline"] = top

    # Timeline
    ws = rd.sheet("Timeline")
    rd.startswith(ws, 1, 1, "Legislative Timeline")
    tc = rd.cols(ws, 7, {"date": "Date", "event": "Event", "actor": "Actor",
                         "src": "Source", "gap": "Days since prior event",
                         "rel": "Days relative to FY2026 start",
                         "after": "Days after the line-item request"})
    tl = []
    r = 8
    while r <= ws.max_row:
        d = ws.cell(r, tc["date"]).value
        if d is None:
            break
        tl.append({"date": txt(d),
                   "event": txt(ws.cell(r, tc["event"]).value),
                   "actor": txt(ws.cell(r, tc["actor"]).value),
                   "src": txt(ws.cell(r, tc["src"]).value),
                   "rel": r2(num(ws.cell(r, tc["rel"]).value))})
        r += 1
    out["timeline"] = tl

    # Glossary
    ws = rd.sheet("Definitions")
    rd.startswith(ws, 1, 1, "Definitions and Field Dictionary")
    rd.startswith(ws, 4, 1, "A. Glossary")
    out["glossary"] = read_simple(rd, "Definitions", 5,
                                  {"term": "Term", "meaning": "Meaning"}, 6)

    # Sources
    ws = rd.sheet("Sources")
    rd.startswith(ws, 1, 1, "Sources")
    out["sources26"] = read_simple(rd, "Sources", 4, {
        "id": "ID", "title": "Title", "author": "Author or publisher",
        "date": "Date", "supplies": "What it supplies", "link": "Link"}, 5)
    # The FY2027 sheet carries the same six columns as the FY2026 one and
    # continues its reference numbering, so the page can present one list.
    ws = rd.sheet("Sources_FY2027")
    rd.startswith(ws, 1, 1, "Sources, FY2027 Extension")
    out["sources27"] = read_simple(rd, "Sources_FY2027", 4, {
        "id": "ID", "title": "Title", "author": "Author or publisher",
        "date": "Date", "supplies": "What it supplies", "link": "Link"}, 5)
    ids = [r["id"] for r in out["sources26"]] + [r["id"] for r in out["sources27"]]
    if len(set(ids)) != len(ids):
        dupes = sorted({i for i in ids if ids.count(i) > 1})
        raise WorkbookMismatch(
            "the two Sources sheets reuse reference numbers: %s. The FY2027 sheet "
            "is meant to continue the numbering, not restart it."
            % ", ".join(dupes))

    # QA
    ws = rd.sheet("QA_Checks")
    rd.startswith(ws, 1, 1, "Quality Assurance and Source Reconciliation")
    out["qa26"] = read_simple(rd, "QA_Checks", 7, {
        "cat": "Category", "measure": "Measure", "book": "This workbook",
        "crs": "CRS narrative", "diff": "Difference", "status": "Status",
        "where": "CRS location", "why": "Explanation"}, 8,
        textkeys=("cat", "measure", "status", "where", "why"))
    # Section B carries a second header row with different columns, so section A
    # is read only as far as the blank row that separates them.
    rd.startswith(ws, 65, 1, "B. Structural checks")
    rd.label(ws, 66, 1, "Check")
    rd.label(ws, 66, 3, "Value")
    rd.label(ws, 66, 4, "Expected")
    rd.label(ws, 66, 5, "Status")
    rd.label(ws, 66, 6, "Comment")
    struct = []
    r = 67
    while r <= ws.max_row:
        k = txt(ws.cell(r, 1).value)
        if not k:
            break
        struct.append({"check": k,
                       "value": r2(num(ws.cell(r, 3).value)),
                       "expect": r2(num(ws.cell(r, 4).value)),
                       "status": txt(ws.cell(r, 5).value),
                       "note": txt(ws.cell(r, 6).value)})
        r += 1
    out["qa26struct"] = struct
    sbad = [q for q in struct if q["status"] not in ("PASS", "")]
    if sbad:
        raise WorkbookMismatch(
            "QA_Checks section B reports %d structural checks that are not PASS, "
            "the first being %r." % (len(sbad), sbad[0]["check"]))
    # QA_FY2027 runs in lettered sections, each with its own header row and its
    # own column labels. The walker takes the first row after each section title
    # as that section's header, verifies it, and reads the rows beneath it. That
    # keeps every check rather than stopping at the first blank line, and keeps
    # each check under the heading it belongs to.
    ws = rd.sheet("QA_FY2027")
    rd.startswith(ws, 1, 1, "QA and Tie-Outs, FY2027 Extension")
    HEADS = ("check", "category", "contradiction", "finding")
    qa27, sections27, cur, want_head = [], [], None, False
    for r in range(5, ws.max_row + 1):
        a = txt(ws.cell(r, 1).value)
        if not a:
            continue
        if re.match(r"^[A-Z]\. ", a):
            cur = {"title": a, "head": [], "rows": [], "notes": []}
            sections27.append(cur)
            want_head = True
            continue
        if cur is None:
            continue
        if want_head:
            want_head = False
            if _norm(a) in HEADS:
                rd.checks += 1
                cur["head"] = [txt(ws.cell(r, c).value) for c in range(1, 7)]
                continue
            # A section of plain prose, such as the closing note on what changed
            # in the FY2026 sheets, carries no header row. Fall through and read
            # its lines as notes.
            cur["prose"] = True
        status = txt(ws.cell(r, 5).value)
        book = ws.cell(r, 2).value
        row = {"check": a,
               "book": r2(num(book)) if isinstance(book, (int, float)) else txt(book),
               "expect": r2(num(ws.cell(r, 3).value)),
               "diff": r2(num(ws.cell(r, 4).value)),
               "status": status,
               "note": txt(ws.cell(r, 6).value),
               "section": cur["title"]}
        if status:
            qa27.append(row)
            cur["rows"].append(row)
        else:
            cur["notes"].append(row)
    out["qa27"] = qa27
    out["qa27sections"] = sections27
    ok = len([q for q in qa27 if q["status"] == "OK"])
    chk = len([q for q in qa27 if q["status"] == "CHECK"])
    other = [q for q in qa27 if q["status"] not in ("OK", "CHECK")]
    if other:
        raise WorkbookMismatch(
            "QA_FY2027 carries %d checks whose status is neither OK nor CHECK, "
            "the first being %r (%r)."
            % (len(other), other[0]["check"], other[0]["status"]))
    if ok < 40 or chk != 1:
        raise WorkbookMismatch(
            "QA_FY2027 should report at least 40 OK checks and exactly one "
            "CHECK, the documented disagreement between the CRS summary and the "
            "CRS tables on the SASC aggregate. It reports %d OK and %d CHECK."
            % (ok, chk))
    out["qa27counts"] = {"ok": ok, "check": chk}
    # PASS and REVIEW are both expected. REVIEW marks a place where CRS's own
    # narrative disagrees with CRS's own tables; the workbook shows both figures
    # rather than reconciling them, and so does the page. Anything else is a
    # failure and stops the build.
    bad = [q for q in out["qa26"]
           if q.get("status") not in ("PASS", "REVIEW", "", None)]
    if bad:
        raise WorkbookMismatch(
            "QA_Checks reports %d rows that are neither PASS nor REVIEW, the "
            "first being %r. The workbook does not tie out, so the page is not "
            "published." % (len(bad), bad[0].get("measure")))

    # Projection
    ws = rd.sheet("Projection")
    rd.startswith(ws, 1, 1, "Projection: the Reconciliation Cliff")
    rd.startswith(ws, 4, 1, "A. Assumptions")
    rd.startswith(ws, 5, 1, "Annual price growth (inflation)")
    rd.startswith(ws, 6, 1, "Annual real growth in program content")
    rd.startswith(ws, 7, 1, "Share of FY2026 mandatory funding replaced")
    rd.startswith(ws, 8, 1, "Unallocated P.L. 119-21 balance")
    out["proj"] = {
        "infl": r2(num(ws.cell(5, 3).value)),
        "real": r2(num(ws.cell(6, 3).value)),
        "repl": r2(num(ws.cell(7, 3).value)),
        "balance": r2(num(ws.cell(8, 3).value)),
        "note": txt(ws.cell(2, 1).value),
    }
    rd.startswith(ws, 11, 1, "B. Flat-line sustainment gap by category")
    pc = rd.cols(ws, 12, {
        "cat": "Category", "app": "FY2026 appropriation",
        "spend": "FY2026 spend plan", "tot": "FY2026 enacted total",
        "dep": "Mandatory dependence", "cum": "Cumulative FY2027 to FY2031"})
    prows = []
    for i, name in enumerate(CATS_ORDER):
        r = 13 + i
        rd.label(ws, r, pc["cat"], name)
        prows.append(dict((k, txt(ws.cell(r, col).value) if k == "cat"
                           else r2(num(ws.cell(r, col).value)))
                          for k, col in pc.items()))
    out["projRows"] = prows

    # Reconciliation, Table A-1 line by line
    ws = rd.sheet("Reconciliation")
    rd.startswith(ws, 1, 1, "Reconciliation: Request versus DOD Spend Plan")
    rd.startswith(ws, 31, 1, "C. Table A-1 line by line")
    out["recon"] = read_simple(rd, "Reconciliation", 32, {
        "n": "#", "cat": "Category", "sys": "System",
        "req": "Recon. request (A-1)", "spend": "FY2026 spend plan",
        "chg": "Change ($B)", "pct": "Change (%)", "dir": "Direction",
        "app": "Discretionary appropriation", "tot": "Enacted total",
        "dep": "Mandatory dependence", "note": "Note"}, 33,
        textkeys=("cat", "sys", "dir", "note"))

    # Personnel_Pay assumptions, kept so the FY2026 pay view can be reproduced
    ws = rd.sheet("Personnel_Pay")
    rd.startswith(ws, 4, 1, "A. Assumptions")
    rd.label(ws, 5, 1, "Assumption")
    rd.label(ws, 5, 3, "Value")
    pp = {}
    for r in range(6, 20):
        k = txt(ws.cell(r, 1).value)
        if not k:
            continue
        v = ws.cell(r, 3).value
        pp[k] = r2(num(v)) if isinstance(v, (int, float)) else txt(v)
    out["personnel"] = pp
    return out

# ---------------------------------------------------------------------------
# Payload
# ---------------------------------------------------------------------------

SERVICE_ORDER = ["Air Force", "Army", "Navy", "Marine Corps", "Space Force",
                 "MDA", "SOCOM", "Multi-service"]


def build_payload(path):
    rd = Reader(path)
    rd.precheck()
    fy26 = read_fy26(rd)
    fy27 = read_fy27(rd)
    cats26 = read_cats26(rd)
    cats27 = read_cats27(rd)
    cmp_ = read_compare(rd)
    pay = read_pay(rd)
    ptab = read_pay_table(rd)
    force = read_force(rd)
    comp = read_composite(rd)
    allow = read_allow(rd)
    ctx = read_context(rd)

    svcs = [s for s in SERVICE_ORDER if any(d["svc"] == s for d in fy26)]
    for d in fy26:
        if d["svc"] and d["svc"] not in svcs:
            svcs.append(d["svc"])

    # The FY2027 sheet carries no analyst service field, so the page inherits one
    # through the crosswalk where a line exists in both years and leaves it blank
    # otherwise rather than guessing.
    svc27 = {}
    by26 = dict((d["id"], d) for d in fy26)
    for row in cmp_["rows"]:
        if row["id26"] and row["id27"] and row["id26"] in by26:
            svc27[row["id27"]] = by26[row["id26"]]["svc"]
    for d in fy27:
        d["svc"] = svc27.get(d["id"], "")

    # Representative years of service, carried from the pay module so the page
    # starts where the workbook starts.
    repyos = dict((g["g"], g["yos"]) for g in pay["grades"])
    payrates = {}
    for g in pay["grades"]:
        row = ptab["table"].get(g["g"])
        if row is None:
            raise WorkbookMismatch(
                "Pay_Table_2026 has no row for pay grade %r, which Pay_Module "
                "prices." % g["g"])
        payrates[g["g"]] = row

    payload = {
        "meta": {
            "version": VERSION,
            "built": datetime.datetime.now().strftime("%d %B %Y"),
            "workbook": os.path.basename(path),
            "checks": rd.checks,
        },
        "cats": CATS_ORDER,
        "services": svcs,
        "fy26": fy26,
        "fy27": fy27,
        "cats26": cats26,
        "cats27": cats27,
        "compare": cmp_,
        "pay": {
            "m1": pay["m1"],
            "assumptions": pay["assumptions"],
            "grades": pay["grades"],
            "weighted": pay["weighted"],
            "steps": ptab["steps"],
            "table": payrates,
            "repyos": repyos,
            "force": force["total"],
            "forcesvc": force["bysvc"],
            "forceOrder": force["order"],
            "forceServices": force["services"],
            "forceSum": force["sum"],
            "composite": comp,
            "allow": allow,
        },
        "ctx": ctx,
    }
    return payload, rd

# ---------------------------------------------------------------------------
# The page: head and stylesheet
# ---------------------------------------------------------------------------

TEMPLATE_HEAD = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="color-scheme" content="light dark">
<title>__TITLE__</title>
<meta name="description" content="__DESC__">
<meta name="author" content="Alyssa I. Agard">
<meta name="copyright" content="Copyright __YEAR__ Alyssa I. Agard. Licensed CC BY 4.0.">
<meta name="license" content="https://creativecommons.org/licenses/by/4.0/">
<meta property="og:title" content="__TITLE__">
<meta property="og:description" content="__DESC__">
<meta property="og:type" content="article">
<meta property="article:author" content="Alyssa I. Agard">
<meta name="theme-color" content="#fde9f5" media="(prefers-color-scheme: light)">
<meta name="theme-color" content="#171015" media="(prefers-color-scheme: dark)">
__ICONS__
<style>
:root{
  --paper:#fffcfd; --panel:#ffffff; --ink:#1f1218; --muted:#8a6f79; --soft:#b9a3ab;
  --rose:#a8506e; --rose2:#d98ba3; --rose3:#f2b8c6; --mauve:#c9b6d9;
  --blush:#fcf1f5; --blush2:#f7e4eb; --line:#ecd9e0; --hair:#f5eaef;
  --good:#2f7d5b; --bad:#b0304a; --warn:#9a6b1f; --navy:#3b3354; --teal:#4a7f92;
  --serif:"Times New Roman",Times,serif;
  --sans:"Segoe UI",Calibri,"Helvetica Neue",Helvetica,Arial,sans-serif;
  --mono:ui-monospace,SFMono-Regular,Menlo,Consolas,"Courier New",monospace;
  --r-card:10px; --r-input:6px; --r-chart:8px;
  --ease:cubic-bezier(.22,.61,.36,1);
  --ease-soft:cubic-bezier(.33,.9,.28,1);
  --t-fast:220ms; --t-mid:420ms; --t-slow:680ms;
  box-sizing:border-box;
}
@media (prefers-color-scheme: dark){
  :root:not([data-theme="light"]){
    --paper:#171015; --panel:#1e151c; --ink:#f7e8ee; --muted:#b394a1; --soft:#7d6672;
    --rose:#e9a3bb; --rose2:#c9738b; --rose3:#8d4460; --mauve:#6d5f80;
    --blush:#241922; --blush2:#2c1f28; --line:#3a2a33; --hair:#281d25;
    --good:#7fc3a1; --bad:#e3899b; --warn:#d8b169; --navy:#463c66; --teal:#7fb0c2;
  }
}
:root[data-theme="dark"]{
  --paper:#171015; --panel:#1e151c; --ink:#f7e8ee; --muted:#b394a1; --soft:#7d6672;
  --rose:#e9a3bb; --rose2:#c9738b; --rose3:#8d4460; --mauve:#6d5f80;
  --blush:#241922; --blush2:#2c1f28; --line:#3a2a33; --hair:#281d25;
  --good:#7fc3a1; --bad:#e3899b; --warn:#d8b169; --navy:#463c66; --teal:#7fb0c2;
}
*{box-sizing:border-box}
html{scroll-behavior:smooth;scroll-padding-top:96px;-webkit-text-size-adjust:100%}
html,body{margin:0;padding:0}
body{background:var(--paper);color:var(--ink);font-family:var(--serif);font-size:16.5px;
  line-height:1.56;-webkit-font-smoothing:antialiased;text-rendering:optimizeLegibility;
  padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
.wrap{max-width:1280px;margin:0 auto;padding:0 40px}
@media(max-width:760px){.wrap{padding:0 16px}}

/* ---- chrome ---- */
.utility{border-bottom:1px solid var(--hair);background:var(--panel)}
.utilityinner{display:flex;justify-content:space-between;align-items:center;gap:14px;
  min-height:34px;font-size:12px;color:var(--soft);flex-wrap:wrap;padding:5px 0;
  font-family:var(--sans)}
.utilityinner span{white-space:nowrap}
tbody td.id,thead th.id{white-space:nowrap;font-family:var(--mono);font-size:12.5px;color:var(--muted)}
.brandrow{display:flex;align-items:center;gap:18px;padding:22px 0 16px;background:var(--panel);
  flex-wrap:wrap}
.tile{width:58px;height:40px;flex:0 0 auto;border:0;outline:0;box-shadow:none;background:none;
  padding:0;line-height:0;border-radius:2px;overflow:hidden}
.tile svg{display:block;border:0;outline:0}
.brandtext{flex:1 1 260px;min-width:0}
.brandtext h1{margin:0;font-size:25px;font-weight:400;letter-spacing:-.01em;line-height:1.16}
.brandtext p{margin:3px 0 0;font-size:13.5px;color:var(--muted);font-style:italic}
.brandtext p.author{font-family:var(--sans);font-style:normal;font-size:12.5px;
  color:var(--soft);margin-top:5px;letter-spacing:.01em}
.amtrow{display:flex;gap:8px}
.amtrow input{flex:1 1 auto;min-width:0}
.amtrow select{flex:0 0 auto;width:auto;min-width:104px}
.amtnote{margin:6px 0 0;font-family:var(--mono);font-size:12px;color:var(--muted);
  font-style:normal;min-height:1em;transition:color var(--t-fast) var(--ease)}
.brandtools{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
.yearswitch{display:inline-flex;border:1px solid var(--line);border-radius:var(--r-input);
  overflow:hidden;background:var(--paper)}
.yearswitch button{appearance:none;border:0;background:transparent;font-family:var(--sans);
  font-size:13px;color:var(--muted);padding:7px 15px;cursor:pointer;
  transition:background .28s var(--ease),color .28s var(--ease)}
.yearswitch button + button{border-left:1px solid var(--line)}
.yearswitch button:hover{background:var(--blush);color:var(--ink)}
.yearswitch button[aria-pressed="true"]{background:var(--rose);color:#fff}
.tool{appearance:none;border:1px solid var(--line);background:var(--paper);color:var(--muted);
  font-family:var(--sans);font-size:12.5px;padding:7px 12px;border-radius:var(--r-input);
  cursor:pointer;transition:border-color .28s var(--ease),color .28s var(--ease),
  background .28s var(--ease)}
.tool:hover{border-color:var(--rose2);color:var(--rose)}
.navrow{position:sticky;top:0;z-index:40;background:var(--panel);border-top:1px solid var(--hair);
  border-bottom:1px solid var(--line)}
.navinner{display:flex;gap:26px;overflow-x:auto;scrollbar-width:none;-ms-overflow-style:none}
.navinner::-webkit-scrollbar{display:none}
.navlink{position:relative;appearance:none;border:0;background:none;font-family:var(--sans);
  font-size:14px;letter-spacing:.005em;color:var(--muted);padding:14px 0 13px;cursor:pointer;
  white-space:nowrap;transition:color .28s var(--ease)}
.navlink::after{content:"";position:absolute;left:0;right:0;bottom:-1px;height:2px;
  background:var(--rose);transform:scaleX(0);transform-origin:left center;
  transition:transform 220ms var(--ease)}
.navlink:hover{color:var(--ink)}
.navlink[aria-selected="true"]{color:var(--ink)}
.navlink[aria-selected="true"]::after{transform:scaleX(1)}

/* ---- hero ---- */
.hero{display:grid;grid-template-columns:1.08fr .92fr;gap:54px;align-items:center;
  padding:56px 0 46px;border-bottom:1px solid var(--hair)}
.kicker{font-size:12px;letter-spacing:.09em;text-transform:uppercase;color:var(--rose);
  font-family:var(--mono);margin:0 0 14px}
.hero h2{margin:0 0 18px;font-size:clamp(34px,5vw,58px);font-weight:400;line-height:1.04;
  letter-spacing:-.022em}
.lead{margin:0;font-size:17.5px;color:var(--muted);line-height:1.62;max-width:52ch}
.herovis{aspect-ratio:4/3;background:var(--panel);border:1px solid var(--line);
  border-radius:var(--r-chart);padding:16px;display:flex;flex-direction:column;min-width:0}
.herovis .vt{font-size:12.5px;color:var(--muted);font-style:italic;margin:0 0 6px}
.herovis .vbody{flex:1;min-height:0;display:flex;align-items:center}
.herovis svg{width:100%;height:auto;display:block}

/* ---- editorial figures, not metric cards ---- */
.figrow{display:grid;grid-template-columns:repeat(4,1fr);gap:0;margin:34px 0 8px}
.fig{padding:2px 22px;border-left:1px solid var(--line)}
.fig:first-child{padding-left:0;border-left:0}
.fig .v{font-family:var(--serif);font-size:31px;line-height:1.12;letter-spacing:-.015em;
  font-variant-numeric:tabular-nums}
/* A figure whose value is words rather than a number needs to sit quieter, or a
   long label swamps the row it belongs to. */
.fig .v.words{font-size:20px;line-height:1.24;letter-spacing:-.008em}
.fig .k{font-size:12.5px;color:var(--muted);font-style:italic;margin-top:4px;line-height:1.35}
.fig .s{font-family:var(--mono);font-size:11.5px;color:var(--soft);margin-top:3px}

/* ---- sections ---- */
.tabpanel{display:none;padding:38px 0 60px}
.tabpanel.on{display:block;animation:panelIn var(--t-mid) var(--ease-soft) both}
@keyframes panelIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
#heroSect{transition:opacity var(--t-mid) var(--ease-soft)}
/* Figures and table bodies are rebuilt wholesale on every recalculation, so the
   swap is softened rather than left to snap. */
.figrow .fig{animation:figIn var(--t-mid) var(--ease-soft) both}
.figrow .fig:nth-child(2){animation-delay:60ms}
.figrow .fig:nth-child(3){animation-delay:120ms}
.figrow .fig:nth-child(4){animation-delay:180ms}
@keyframes figIn{from{opacity:0;transform:translateY(6px)}to{opacity:1;transform:none}}
.fig .v{transition:color var(--t-fast) var(--ease)}
.ledger .li{animation:figIn var(--t-fast) var(--ease-soft) both}
.sect{padding:34px 0 0}
.sect h3{margin:0 0 12px;font-size:29px;font-weight:400;letter-spacing:-.018em;line-height:1.14}
.sect h4{margin:30px 0 10px;font-size:19px;font-weight:400;letter-spacing:-.01em}
.sect p{margin:0 0 14px;color:var(--muted);max-width:74ch;line-height:1.62}
.sect p.body{color:var(--ink)}
.note{font-size:13.5px;color:var(--soft);font-style:italic;max-width:78ch;margin:10px 0 0}
.rule{border:0;border-top:1px solid var(--hair);margin:34px 0}

/* ---- feature cards ---- */
.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:22px;margin:30px 0 4px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:var(--r-card);
  padding:22px 22px 20px;transition:border-color var(--t-mid) var(--ease-soft),
  transform var(--t-mid) var(--ease-soft),box-shadow var(--t-mid) var(--ease-soft)}
.card:hover{border-color:var(--rose2);transform:translateY(-3px);
  box-shadow:0 10px 30px -18px var(--rose2)}
.card .ic svg{transition:transform var(--t-mid) var(--ease-soft)}
.card:hover .ic svg{transform:translateY(-1px) scale(1.04)}
.card .ic{color:var(--rose);margin-bottom:12px;line-height:0}
.card h5{margin:0 0 7px;font-size:17px;font-weight:400}
.card p{margin:0;font-size:14px;color:var(--muted);line-height:1.55}

/* ---- panels, controls ---- */
.panel{background:var(--panel);border:1px solid var(--line);border-radius:var(--r-card);
  padding:22px;margin:22px 0}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:22px}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:18px}
.field{margin:0 0 14px}
.field label{display:block;font-size:12.5px;color:var(--muted);margin-bottom:5px;
  font-family:var(--sans)}
input[type=text],input[type=number],select,textarea{width:100%;font-family:var(--serif);
  font-size:14.5px;color:var(--ink);background:var(--paper);border:1px solid var(--line);
  border-radius:var(--r-input);padding:8px 10px;
  transition:border-color var(--t-fast) var(--ease),box-shadow var(--t-fast) var(--ease)}
input[type=number],.mono,td.num,th.num{font-family:var(--mono);font-variant-numeric:tabular-nums}
input:focus,select:focus,textarea:focus{outline:0;border-color:var(--rose2);
  box-shadow:0 0 0 3px var(--blush2)}
.rangerow .rv,.field label .mono{transition:color var(--t-fast) var(--ease)}
textarea{min-height:132px;line-height:1.6;resize:vertical}
input[type=range]{width:100%;accent-color:var(--rose)}
.rangerow{display:flex;align-items:center;gap:12px}
.rangerow .rv{font-family:var(--mono);font-size:13px;color:var(--ink);min-width:74px;text-align:right}
.btn{appearance:none;border:1px solid var(--rose);background:var(--rose);color:#fff;
  font-family:var(--sans);font-size:13.5px;padding:9px 17px;border-radius:var(--r-input);
  cursor:pointer;transition:opacity .28s var(--ease),transform .28s var(--ease),
  box-shadow .28s var(--ease)}
.btn:hover{opacity:.9;transform:translateY(-1px);box-shadow:0 3px 10px -4px var(--rose)}
.btn:active{transform:translateY(0);box-shadow:none}
.btn.ghost{background:transparent;color:var(--rose)}
.btn.ghost:hover{background:var(--blush)}
.btnrow{display:flex;gap:10px;flex-wrap:wrap;margin-top:12px}

/* ---- tables ---- */
.tablewrap{overflow-x:auto;margin:16px 0 0;-webkit-overflow-scrolling:touch}
table{width:100%;border-collapse:collapse;font-size:14px}
thead th{font-size:12.5px;font-weight:400;color:var(--muted);text-transform:none;letter-spacing:0;
  background:transparent;text-align:left;padding:9px 11px;border-bottom:1px solid var(--line);
  white-space:nowrap;position:relative;font-family:var(--sans);
  transition:color .25s var(--ease)}
thead th.sortable{cursor:pointer;user-select:none}
thead th.sortable:hover{color:var(--rose)}
thead th .arrow{font-family:var(--mono);font-size:10px;color:var(--rose);margin-left:4px}
tbody td{padding:7px 11px;border-bottom:1px solid var(--hair);vertical-align:top}
tbody tr{transition:background var(--t-fast) var(--ease)}
tbody tr:hover{background:var(--blush)}
tbody{animation:sliceIn var(--t-mid) var(--ease-soft) both}
td.num,th.num{text-align:right;font-size:13.5px}
tfoot td{padding:9px 11px;border-top:1px solid var(--line);font-weight:400;background:var(--blush)}
tbody td.name{min-width:190px}
.pos{color:var(--good)}
.neg{color:var(--bad)}
.dim{color:var(--soft)}
.pending{color:var(--soft);font-style:italic;font-family:var(--serif);font-size:14px}
.tag{display:inline-block;font-family:var(--sans);font-size:11px;padding:1px 7px;
  border-radius:3px;background:var(--blush2);color:var(--muted);white-space:nowrap}
.tag.ok{color:var(--good)}
.tag.chk{color:var(--warn)}
.tag.rev{color:var(--warn)}

/* ---- charts ---- */
.chart{background:var(--panel);border:1px solid var(--line);border-radius:var(--r-chart);
  padding:16px;margin:18px 0}
.chart .ct{font-size:13px;color:var(--muted);font-style:italic;margin:0 0 10px}
svg.svg{width:100%;height:auto;display:block;overflow:visible}
/* Bars and slices grow from their baseline once, so a recalculation reads as a
   movement rather than a jump cut. */
svg.svg rect,svg.svg path{transition:fill var(--t-fast) var(--ease)}
svg.svg.anim rect{animation:barIn var(--t-slow) var(--ease-soft) both;transform-box:fill-box;
  transform-origin:bottom center}
svg.svg.anim.h rect{transform-origin:left center}
@keyframes barIn{from{transform:scaleY(.001)}to{transform:scaleY(1)}}
svg.svg.anim.h rect{animation-name:barInH}
@keyframes barInH{from{transform:scaleX(.001)}to{transform:scaleX(1)}}
svg.svg.anim path{animation:sliceIn var(--t-slow) var(--ease-soft) both}
@keyframes sliceIn{from{opacity:0}to{opacity:1}}
svg.svg.anim text,svg.svg.anim line{animation:sliceIn var(--t-mid) var(--ease-soft) both;
  animation-delay:180ms}
svg.svg rect:hover,svg.svg path:hover{filter:brightness(1.07)}
svg .grid{stroke:var(--hair);stroke-width:1}
svg .ax{fill:var(--muted);font-family:var(--mono);font-size:10.5px}
svg .axb{fill:var(--ink);font-family:var(--serif);font-size:12px}
svg .zero{stroke:var(--line);stroke-width:1}

/* ---- ledger, reveal ---- */
.ledger{margin:14px 0 0;font-size:14px}
.ledger .li{padding:9px 0;border-bottom:1px solid var(--hair)}
.ledger .li .rd{color:var(--ink)}
.ledger .li .mt{color:var(--soft);font-size:12.5px;font-style:italic;margin-top:2px}
.ledger .li.bad .rd{color:var(--bad)}
.reveal{opacity:0;transform:translateY(12px);
  transition:opacity var(--t-slow) var(--ease-soft),transform var(--t-slow) var(--ease-soft)}
.reveal.in{opacity:1;transform:none}
.reveal:nth-of-type(2){transition-delay:70ms}
.reveal:nth-of-type(3){transition-delay:140ms}
.panel{transition:border-color var(--t-mid) var(--ease-soft)}
a{color:var(--rose);text-decoration:none;border-bottom:1px solid transparent;
  transition:border-color var(--t-fast) var(--ease),color var(--t-fast) var(--ease)}
a:hover{border-bottom-color:var(--rose2)}
.sronly{position:absolute;width:1px;height:1px;padding:0;margin:-1px;overflow:hidden;
  clip:rect(0 0 0 0);white-space:nowrap;border:0}
footer{border-top:1px solid var(--line);margin-top:40px;padding:30px 0 52px;font-size:13px;
  color:var(--soft)}
footer p{margin:0 0 8px;max-width:80ch}
.byline{display:grid;grid-template-columns:1.1fr 1fr;gap:26px;align-items:start;
  padding-bottom:22px;margin-bottom:20px;border-bottom:1px solid var(--hair)}
.byline .who{font-family:var(--serif);font-size:18px;color:var(--ink);margin:0 0 4px}
.byline .role{font-size:12.5px;color:var(--muted);font-style:italic;margin:0 0 10px}
.byline ul{list-style:none;margin:0;padding:0;font-family:var(--sans);font-size:12.5px}
.byline li{margin:0 0 4px}
.licence{display:flex;gap:14px;align-items:flex-start}
.licence .cc{flex:0 0 auto;line-height:0;margin-top:2px}
.licence p{margin:0;font-size:12.5px;line-height:1.55}
.licence strong{font-weight:400;color:var(--muted)}
@media(max-width:760px){.byline{grid-template-columns:1fr;gap:18px}}

@media(max-width:1000px){
  .hero{grid-template-columns:1fr;gap:30px;padding:38px 0 32px}
  .cards{grid-template-columns:1fr}
  .grid2,.grid3{grid-template-columns:1fr}
  .figrow{grid-template-columns:1fr 1fr;gap:20px 0}
  .fig{padding:2px 18px}
  .fig:nth-child(odd){padding-left:0;border-left:0}
}
@media(max-width:760px){
  body{font-size:16px}
  .utilityinner span{white-space:normal}
  .utilityinner{font-size:12px}
  .brandtext h1{font-size:21px}
  .sect h3{font-size:24px}
  .figrow{grid-template-columns:1fr}
  .fig{padding:0;border-left:0;border-top:1px solid var(--line);padding-top:12px;margin-top:2px}
  .fig:first-child{border-top:0;padding-top:0}
  table{font-size:13px}
}
@media (prefers-reduced-motion: reduce){
  html{scroll-behavior:auto}
  *,*::before,*::after{animation-duration:.001ms !important;animation-delay:0ms !important;
    animation-iteration-count:1 !important;transition-duration:.001ms !important;
    transition-delay:0ms !important}
  .reveal{opacity:1;transform:none}
  .card:hover{transform:none;box-shadow:none}
  .btn:hover{transform:none;box-shadow:none}
  .navlink::after{transition:none}
  .tabpanel.on,.figrow .fig,.ledger .li,tbody,svg.svg.anim rect,svg.svg.anim path,
  svg.svg.anim text,svg.svg.anim line{animation:none}
}
@media print{
  .navrow,.brandtools,.utility,.btnrow,.tool{display:none}
  .tabpanel{display:block !important;page-break-before:always}
  body{background:#fff;color:#000;font-size:11pt}
  .panel,.card,.chart,.herovis{border:1px solid #ccc;background:#fff}
  a{color:#000;text-decoration:none}
}
</style>
</head>
<body>
"""

TEMPLATE_BODY = r"""
<div class="utility"><div class="wrap utilityinner">
  <span>Congressional Research Service R48860 and R49023</span>
  <span>Alyssa I. Agard &middot; Built __BUILT__ from __WORKBOOK__ &middot; CC BY 4.0</span>
</div></div>

<div class="brandrow wrap">
  <span class="tile" aria-hidden="true">
    <svg viewBox="0 0 87 60" width="58" height="40" role="img">
      <rect width="87" height="60" fill="#f4f2f4"/>
      <g fill="#b3243b">
        <rect y="0" width="87" height="4.615"/><rect y="9.231" width="87" height="4.615"/>
        <rect y="18.462" width="87" height="4.615"/><rect y="27.692" width="87" height="4.615"/>
        <rect y="36.923" width="87" height="4.615"/><rect y="46.154" width="87" height="4.615"/>
        <rect y="55.385" width="87" height="4.615"/>
      </g>
      <rect width="38" height="32.3" fill="#2a3560"/>
      <g fill="#f4f2f4">
        <circle cx="6" cy="6" r="1.7"/><circle cx="14" cy="6" r="1.7"/><circle cx="22" cy="6" r="1.7"/>
        <circle cx="30" cy="6" r="1.7"/><circle cx="10" cy="13" r="1.7"/><circle cx="18" cy="13" r="1.7"/>
        <circle cx="26" cy="13" r="1.7"/><circle cx="6" cy="20" r="1.7"/><circle cx="14" cy="20" r="1.7"/>
        <circle cx="22" cy="20" r="1.7"/><circle cx="30" cy="20" r="1.7"/><circle cx="10" cy="27" r="1.7"/>
        <circle cx="18" cy="27" r="1.7"/><circle cx="26" cy="27" r="1.7"/>
      </g>
    </svg>
  </span>
  <div class="brandtext">
    <h1>Defense Budget Lab, FY2026 and FY2027</h1>
    <p>Selected weapon systems, personnel pay, and what a change to one would cost</p>
    <p class="author">By Alyssa I. Agard</p>
  </div>
  <div class="brandtools">
    <div class="yearswitch" role="group" aria-label="Fiscal year">
      <button type="button" id="y26" aria-pressed="false">FY2026</button>
      <button type="button" id="y27" aria-pressed="true">FY2027</button>
      <button type="button" id="yboth" aria-pressed="false">Both</button>
    </div>
    <button type="button" class="tool" id="themeBtn">Theme</button>
    <button type="button" class="tool" id="linkBtn">Copy link</button>
    <button type="button" class="tool" id="memoBtn">Memo</button>
    <button type="button" class="tool" id="xlsxBtn">Workbook</button>
  </div>
</div>

<div class="navrow"><div class="wrap"><div class="navinner" role="tablist" aria-label="Sections">
  <button type="button" class="navlink" role="tab" data-tab="overview"   aria-selected="true">Overview</button>
  <button type="button" class="navlink" role="tab" data-tab="scenario"   aria-selected="false">Scenario composer</button>
  <button type="button" class="navlink" role="tab" data-tab="claim"      aria-selected="false">Claim tester</button>
  <button type="button" class="navlink" role="tab" data-tab="pay"        aria-selected="false">Pay and compensation</button>
  <button type="button" class="navlink" role="tab" data-tab="projection" aria-selected="false">Projection</button>
  <button type="button" class="navlink" role="tab" data-tab="compare"    aria-selected="false">Year on year</button>
  <button type="button" class="navlink" role="tab" data-tab="data"       aria-selected="false">The data</button>
  <button type="button" class="navlink" role="tab" data-tab="method"     aria-selected="false">Method and sources</button>
</div></div></div>

<main class="wrap">

<section class="hero reveal" id="heroSect">
  <div>
    <p class="kicker" id="heroKicker">Fiscal year 2027</p>
    <h2 id="heroTitle">What the request asks for, and what the committees did to it</h2>
    <p class="lead" id="heroLead"></p>
  </div>
  <div class="herovis">
    <p class="vt" id="heroVisT"></p>
    <div class="vbody" id="heroVis"></div>
  </div>
</section>

<div class="figrow reveal" id="heroFigs"></div>

<div id="t-overview" class="tabpanel on" role="tabpanel">
  <div class="cards reveal">
    <div class="card">
      <div class="ic"><svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3"><path d="M3 20h18M6 20V9M11 20V4M16 20v-7M21 20v-11"/></svg></div>
      <h5>Every figure traces to a page</h5>
      <p>Each line carries the CRS table it came from, the page it was printed on, and the footnote attached to it. A blank means CRS printed nothing there. It is never shown as a zero.</p>
    </div>
    <div class="card">
      <div class="ic"><svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3"><path d="M4 7h16M4 12h16M4 17h10"/><circle cx="18.5" cy="17" r="2.5"/></svg></div>
      <h5>Written instructions, priced</h5>
      <p>Type an instruction the way you would say it to a staffer. The composer reads it, prices it against the baseline you choose, and shows the offset it implies.</p>
    </div>
    <div class="card">
      <div class="ic"><svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.3"><path d="M12 3v18M5 8l7-5 7 5M5 16l7 5 7-5"/></svg></div>
      <h5>Two years, one spine</h5>
      <p>The FY2026 and FY2027 reports do not use the same column names. The crosswalk says which pairs are comparable, and where a programme appears in only one year it stays blank rather than reading as a cut.</p>
    </div>
  </div>

  <div class="sect reveal">
    <p class="kicker">By category</p>
    <h3 id="ovTitle">The eight categories CRS tabulates</h3>
    <p id="ovLead"></p>
    <div class="chart"><p class="ct" id="ovChartT"></p><div id="ovChart"></div></div>
    <div class="tablewrap"><table id="ovTable"></table></div>
    <p class="note" id="ovNote"></p>
  </div>

  <div class="sect reveal">
    <p class="kicker">Where the money concentrates</p>
    <h3>The largest lines, and what the committees did</h3>
    <p id="topLead"></p>
    <div class="grid2">
      <div class="chart"><p class="ct" id="topChartT"></p><div id="topChart"></div></div>
      <div class="chart"><p class="ct" id="actChartT"></p><div id="actChart"></div></div>
    </div>
  </div>

  <div class="sect reveal">
    <p class="kicker">Context</p>
    <h3>The selected systems inside the whole request</h3>
    <div class="tablewrap"><table id="ctxTable"></table></div>
  </div>
</div>

<div id="t-scenario" class="tabpanel" role="tabpanel">
  <div class="sect">
    <p class="kicker">Scenario composer</p>
    <h3>Write the change, and see what it costs</h3>
    <p>One instruction a line. The parser understands add, cut, set, and zero, in dollars or in percentages, against a named system, a category, a service, or the whole portfolio. Add an offset clause and the money is taken from somewhere.</p>
    <div class="grid2">
      <div class="panel">
        <div class="field">
          <label for="scenText">Instructions</label>
          <textarea id="scenText" spellcheck="false" placeholder="add 2 billion to Virginia-class Submarine for a third boat&#10;cut 10% from Missile Defense, offset by the rest of the portfolio&#10;zero out the Sentinel"></textarea>
        </div>
        <div class="grid2">
          <div class="field">
            <label for="scenBase">Baseline</label>
            <select id="scenBase"></select>
          </div>
          <div class="field">
            <label for="scenOff">Offset rule</label>
            <select id="scenOff">
              <option value="none">No offset, the topline moves</option>
              <option value="all">Pro rata across untouched lines</option>
              <option value="cat">Hold each category constant</option>
              <option value="sys">One system absorbs the net</option>
            </select>
          </div>
        </div>
        <div class="field" id="scenSysWrap" style="display:none">
          <label for="scenOffSys">Absorbing system</label>
          <select id="scenOffSys"></select>
        </div>
        <div class="field">
          <label for="scenGlobal">Across the board <span class="mono" id="scenGlobalV">0.0%</span></label>
          <input type="range" id="scenGlobal" min="-25" max="25" step="0.5" value="0">
        </div>
        <div class="btnrow">
          <button type="button" class="btn" id="scenApply">Price it</button>
          <button type="button" class="btn ghost" id="scenClear">Clear</button>
          <button type="button" class="btn ghost" id="scenCsv">Download CSV</button>
        </div>
      </div>
      <div>
        <div class="figrow" id="scenFigs" style="margin-top:0"></div>
        <div class="ledger" id="ledger"></div>
      </div>
    </div>
    <div class="chart"><p class="ct" id="scenChartT"></p><div id="scenChart"></div></div>
    <h4>Line by line</h4>
    <div class="tablewrap"><table id="scenTable"></table></div>
  </div>
</div>

<div id="t-claim" class="tabpanel" role="tabpanel">
  <div class="sect">
    <p class="kicker">Claim tester</p>
    <h3>Put a number in proportion</h3>
    <p>A figure on its own says little. This takes a change and expresses it four ways: as a share of its own line, as units of the thing it buys, as a cut somewhere else, and as an across-the-board military pay raise. The change can be a percentage of a line CRS printed, an amount you type in any unit from dollars to trillions, or a claim of your own that appears in neither report.</p>
    <div class="grid3">
      <div class="panel">
        <div class="field"><label for="clSys">Line, or a claim of your own</label>
          <select id="clSys"></select></div>
        <div class="field" id="clOwnWrap" style="display:none">
          <label for="clOwnLab">What the money is for</label>
          <input type="text" id="clOwnLab" placeholder="for example, a grant to one researcher"></div>
        <div class="field"><label for="clBase">Measured against</label><select id="clBase"></select></div>
        <div class="field"><label for="clType">Change</label>
          <select id="clType"><option value="pct">A percentage</option><option value="usd">An amount you type</option></select>
        </div>
        <div class="field"><label for="clVal" id="clValLab">Percentage</label>
          <div class="amtrow">
            <input type="number" id="clVal" value="10" step="0.5">
            <select id="clUnit" style="display:none">
              <option value="1000000000">billion</option>
              <option value="1000000">million</option>
              <option value="1000">thousand</option>
              <option value="1">dollars</option>
            </select>
          </div>
          <p class="note amtnote" id="clAmt"></p></div>
        <div class="field"><label for="clOff">Paid for by</label>
          <select id="clOff">
            <option value="none">Nothing, the topline moves</option>
            <option value="sys">One other system</option>
            <option value="cat">The rest of its category</option>
            <option value="all">Every other selected system</option>
          </select></div>
        <div class="field" id="clOffWrap" style="display:none"><label for="clOffSys">Which system</label>
          <select id="clOffSys"></select></div>
        <div class="field"><label for="clRef">Compared with buying</label><select id="clRef"></select></div>
      </div>
      <div style="grid-column:span 2">
        <div id="claimBody"></div>
      </div>
    </div>
  </div>
</div>

<div id="t-pay" class="tabpanel" role="tabpanel">
  <div class="sect">
    <p class="kicker">Pay and compensation</p>
    <h3>What a military pay raise costs, grade by grade</h3>
    <p id="payLead"></p>
    <div class="grid3">
      <div class="panel">
        <div class="field"><label for="payYear">Fiscal year priced</label>
          <select id="payYear"><option value="2026">FY2026</option><option value="2027" selected>FY2027</option></select></div>
        <div class="field"><label for="paySvc">Service scope</label><select id="paySvc"></select></div>
        <div class="field"><label for="payAta">Across the board <span class="mono" id="payAtaV">0.0%</span></label>
          <input type="range" id="payAta" min="0" max="15" step="0.1" value="0"></div>
        <div class="field"><label for="payGrade">Add to one grade</label><select id="payGrade"></select></div>
        <div class="field"><label for="payGradePct">That grade, extra <span class="mono" id="payGradePctV">10.0%</span></label>
          <input type="range" id="payGradePct" min="0" max="30" step="0.5" value="10"></div>
        <div class="field"><label for="payMonth">Effective month (1 is January)</label>
          <input type="number" id="payMonth" min="1" max="12" value="1"></div>
        <div class="field"><label for="payRpa">Include retired pay accrual</label>
          <select id="payRpa"><option value="1">Yes</option><option value="0">No</option></select></div>
        <div class="btnrow"><button type="button" class="btn ghost" id="payReset">Reset</button></div>
      </div>
      <div style="grid-column:span 2">
        <div class="figrow" id="payFigs" style="margin-top:0"></div>
        <div class="chart"><p class="ct" id="payChartT"></p><div id="payChart"></div></div>
      </div>
    </div>
    <h4>Grade by grade</h4>
    <div class="tablewrap"><table id="payTable"></table></div>
    <p class="note" id="payNote"></p>
    <hr class="rule">
    <h4>Regular military compensation</h4>
    <p id="rmcLead"></p>
    <div class="tablewrap"><table id="rmcTable"></table></div>
    <p class="note" id="rmcNote"></p>
    <h4>The model against the published composite rate</h4>
    <div class="tablewrap"><table id="compTable"></table></div>
  </div>
</div>

<div id="t-projection" class="tabpanel" role="tabpanel">
  <div class="sect">
    <p class="kicker">Projection</p>
    <h3>The reconciliation cliff</h3>
    <p id="projLead"></p>
    <div class="grid3">
      <div class="panel">
        <div class="field"><label for="prInfl">Price growth a year <span class="mono" id="prInflV">2.0%</span></label>
          <input type="range" id="prInfl" min="0" max="8" step="0.1" value="2"></div>
        <div class="field"><label for="prReal">Real growth in content <span class="mono" id="prRealV">0.0%</span></label>
          <input type="range" id="prReal" min="-5" max="8" step="0.1" value="0"></div>
        <div class="field"><label for="prRepl">Mandatory money replaced <span class="mono" id="prReplV">0%</span></label>
          <input type="range" id="prRepl" min="0" max="100" step="1" value="0"></div>
        <p class="note">Set the replacement share to 100 percent to model Congress backfilling every reconciliation dollar with discretionary appropriations.</p>
      </div>
      <div style="grid-column:span 2">
        <div class="figrow" id="projFigs" style="margin-top:0"></div>
        <div class="chart"><p class="ct" id="projChartT"></p><div id="projChart"></div></div>
      </div>
    </div>
    <div class="tablewrap"><table id="projTable"></table></div>
    <p class="note" id="projNote"></p>
  </div>
</div>

<div id="t-compare" class="tabpanel" role="tabpanel">
  <div class="sect">
    <p class="kicker">Year on year</p>
    <h3>FY2026 against FY2027, line by line</h3>
    <p>The two reports do not share a vocabulary. What follows is the crosswalk the workbook settled on, stated before any figure is compared.</p>
    <div class="panel"><div id="cmpNotes"></div></div>
    <div class="figrow" id="cmpFigs"></div>
    <div class="chart"><p class="ct" id="cmpChartT"></p><div id="cmpChart"></div></div>
    <h4>The bridge</h4>
    <div class="tablewrap"><table id="bridgeTable"></table></div>
    <h4>Lines that exist in one year only</h4>
    <div class="tablewrap"><table id="onlyTable"></table></div>
    <h4>The joined spine</h4>
    <div class="grid3">
      <div class="field"><label for="cmpQ">Search</label><input type="text" id="cmpQ" placeholder="name or ID"></div>
      <div class="field"><label for="cmpCat">Category</label><select id="cmpCat"></select></div>
      <div class="field"><label for="cmpPres">Present in</label>
        <select id="cmpPres"><option value="all">Either year</option><option value="Both years">Both years</option>
        <option value="FY2027 only">FY2027 only</option><option value="FY2026 only">FY2026 only</option></select></div>
    </div>
    <div class="tablewrap"><table id="cmpTable"></table></div>
  </div>
</div>

<div id="t-data" class="tabpanel" role="tabpanel">
  <div class="sect">
    <p class="kicker">The data</p>
    <h3 id="dTitle">Every line CRS printed</h3>
    <p id="dLead"></p>
    <div class="grid3">
      <div class="field"><label for="dQ">Search</label><input type="text" id="dQ" placeholder="name, ID, or category"></div>
      <div class="field"><label for="dCat">Category</label><select id="dCat"></select></div>
      <div class="field"><label for="dSvc">Service</label><select id="dSvc"></select></div>
    </div>
    <div class="btnrow"><button type="button" class="btn ghost" id="dCsv">Download CSV</button>
      <span class="note" id="dCount" style="margin:0;align-self:center"></span></div>
    <div class="tablewrap"><table id="dTable"></table></div>
    <p class="note">Click a column heading to sort. A blank cell means CRS printed no figure on that line. It is not a zero and it is not a cut.</p>
  </div>
</div>

<div id="t-method" class="tabpanel" role="tabpanel">
  <div class="sect">
    <p class="kicker">Method</p>
    <h3>How this was built, and where it disagrees with itself</h3>
    <p class="body">Every figure on this page is read from the workbook. Nothing is recomputed here, and no figure is adjusted to make a total come out. Where a computed total and a figure CRS states in its own narrative disagree, both are shown and the gap is named.</p>
    <h4>Tie-outs</h4>
    <div class="tablewrap"><table id="qaTable"></table></div>
    <h4>Checks on the FY2027 extension</h4>
    <div class="tablewrap"><table id="qa27Table"></table></div>
    <h4>Contradictions in the sources, left standing</h4>
    <div id="contra"></div>
    <h4>Sources</h4>
    <div class="tablewrap"><table id="srcTable"></table></div>
    <h4>Glossary</h4>
    <div class="tablewrap"><table id="glossTable"></table></div>
    <h4>Legislative timeline</h4>
    <div class="tablewrap"><table id="tlTable"></table></div>
  </div>
</div>

</main>

<footer><div class="wrap">
  <div class="byline">
    <div>
      <p class="who">Alyssa I. Agard</p>
      <p class="role">Author and creator of this page and of the workbook behind it</p>
      <ul>
        <li><a href="https://alyssaagard.dev" rel="noopener noreferrer" target="_blank">alyssaagard.dev</a></li>
        <li><a href="https://www.linkedin.com/in/alyssaagard/" rel="noopener noreferrer" target="_blank">linkedin.com/in/alyssaagard</a></li>
        <li><a href="mailto:economics@alyssaagard.dev">economics@alyssaagard.dev</a></li>
      </ul>
    </div>
    <div class="licence">
      <span class="cc" aria-hidden="true">
        <svg width="62" height="22" viewBox="0 0 62 22" role="img">
          <g fill="none" stroke="currentColor" stroke-width="1.4">
            <circle cx="11" cy="11" r="9.3"/><circle cx="35" cy="11" r="9.3"/>
          </g>
          <text x="11" y="15.4" text-anchor="middle" font-family="Segoe UI,Calibri,Arial,sans-serif"
                font-size="11" fill="currentColor">cc</text>
          <text x="35" y="15.4" text-anchor="middle" font-family="Segoe UI,Calibri,Arial,sans-serif"
                font-size="11" fill="currentColor">by</text>
        </svg>
      </span>
      <p><strong>Licence.</strong> This page, and the workbook it is built from, are released by
      Alyssa I. Agard under a
      <a href="https://creativecommons.org/licenses/by/4.0/" rel="license noopener noreferrer" target="_blank">Creative
      Commons Attribution 4.0 International Licence</a>. You may share and adapt the work, including
      commercially, so long as you give credit, link to the licence, and say whether you changed
      anything. The underlying Congressional Research Service reports are works of the United States
      government and are not subject to copyright; the transcription, the crosswalk, the derived
      figures and this page are the author's own.</p>
    </div>
  </div>
  <p id="footNote"></p>
  <p>Amounts are billions of dollars of budget authority unless a column says otherwise. Figures are as CRS printed them. Service labels on the FY2026 lines are this workbook's classification, not a CRS field, and the FY2027 lines inherit that label only where the crosswalk pairs them with an FY2026 line.</p>
  <p>This page is a reading of Congressional Research Service reports. It is not a government publication and carries no endorsement.</p>
</div></footer>
"""

TEMPLATE_JS = r"""
<script>
"use strict";
const DATA = __DATA__;

/* ---------- small helpers ---------- */
const $ = s => document.querySelector(s);
const $$ = s => Array.prototype.slice.call(document.querySelectorAll(s));
const n = v => (v === null || v === undefined || v === '') ? 0 : +v;
const has = v => v !== null && v !== undefined && v !== '';
const sum = (a, f) => a.reduce((t, x) => t + f(x), 0);
const esc = s => String(s === null || s === undefined ? '' : s)
  .replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const N = (v, d) => isFinite(v)
  ? v.toLocaleString('en-US', {minimumFractionDigits: d || 0, maximumFractionDigits: d || 0}) : 'n/a';
const B = (v, d) => (v < 0 ? '-' : '') + '$' + Math.abs(v).toFixed(d === undefined ? 2 : d) + 'B';
const Bs = (v, d) => (v >= 0 ? '+' : '-') + '$' + Math.abs(v).toFixed(d === undefined ? 2 : d) + 'B';
const P = (v, d) => isFinite(v) ? (v * 100).toFixed(d === undefined ? 1 : d) + '%' : 'n/a';
const Ps = (v, d) => isFinite(v) ? (v >= 0 ? '+' : '') + (v * 100).toFixed(d === undefined ? 1 : d) + '%' : 'n/a';
const USD = v => isFinite(v) ? '$' + N(Math.round(v)) : 'n/a';
const cl = v => v > 0.0005 ? 'pos' : (v < -0.0005 ? 'neg' : 'dim');
/* A blank is a blank. CRS printing nothing is not CRS printing zero. */
const BLANK = '<span class="pending">not printed</span>';
const cellB = (v, f) => has(v) ? f(v) : BLANK;
function money(v) {
  const a = Math.abs(v);
  if (a === 0) return '$0';
  if (a < 0.001) return (v < 0 ? '-' : '') + '$' + N(a * 1000000) + 'K';
  if (a < 1) return (v < 0 ? '-' : '') + '$' + (a * 1000).toFixed(a * 1000 < 10 ? 1 : 0) + 'M';
  return B(v, 2);
}
const moneyS = v => (v >= 0 ? '+' : '') + money(v);

/* ---------- year-aware normalisation ----------
   The two reports name their columns differently. Everything downstream works
   on one shape, so each year is mapped onto it once, here, and the mapping is
   stated rather than assumed. */
const YEARS = {
  2026: {
    label: 'FY2026', report: 'CRS R48860, updated 2 July 2026',
    rows: DATA.fy26, cats: DATA.cats26,
    disc: d => d.base, mand: d => d.recon, total: d => d.total,
    unit: d => d.unit, qty: d => d.baseQ,
    marks: [
      {k: 'hasc',  label: 'H.R. 3838 HASC'},
      {k: 'sasc',  label: 'S. 2296 SASC'},
      {k: 'ndaa',  label: 'Enacted NDAA'},
      {k: 'hac',   label: 'H.R. 4016 HAC'},
      {k: 'sac',   label: 'S. 2572 SAC'},
      {k: 'app',   label: 'Enacted appropriation'}
    ],
    baselines: {
      disc:  {label: 'Base request (discretionary)',           f: d => n(d.base)},
      total: {label: 'Total request (base plus reconciliation)', f: d => n(d.total)},
      ndaa:  {label: 'Enacted NDAA, P.L. 119-60',              f: d => n(d.ndaa)},
      app:   {label: 'Enacted appropriation, P.L. 119-75',     f: d => n(d.app)},
      enact: {label: 'Enacted total (appropriation plus spend plan)', f: d => n(d.app) + n(d.spend)}
    },
    defBase: 'app',
    catDisc: 'base', catMand: 'recon', catTotal: 'total'
  },
  2027: {
    label: 'FY2027', report: 'CRS R49023, version 4, 3 September 2026',
    rows: DATA.fy27, cats: DATA.cats27,
    disc: d => d.disc, mand: d => d.mand, total: d => d.total,
    unit: d => d.unit, qty: d => d.discQ,
    marks: [
      {k: 'house', label: 'H.R. 8800 House'},
      {k: 'sasc',  label: 'S. 4784 SASC'},
      {k: 'hac',   label: 'H.R. 9495 HAC'}
    ],
    baselines: {
      disc:  {label: 'Discretionary request',                  f: d => n(d.disc)},
      total: {label: 'Total request (discretionary plus mandatory)', f: d => n(d.total)},
      house: {label: 'H.R. 8800, as passed the House',         f: d => n(d.house)},
      sasc:  {label: 'S. 4784, as reported by SASC',           f: d => n(d.sasc)},
      hac:   {label: 'H.R. 9495, as reported by HAC',          f: d => n(d.hac)}
    },
    defBase: 'disc',
    catDisc: 'disc', catMand: 'mand', catTotal: 'total'
  }
};
const CATS = DATA.cats;
const SVCS = DATA.services;

let YEAR = 2027;
const Y = () => YEARS[YEAR];
const ROWS = () => Y().rows;
const BASES = () => Y().baselines;

/* Lines that sit outside the eight category tables (the FY2027 appendix) are
   excluded from category arithmetic, exactly as the workbook excludes them. */
const inCat = d => d.inCat === undefined || d.inCat === '' || d.inCat === 'Yes';
const catRows = () => ROWS().filter(inCat);

let byName = {};
function reindex() {
  byName = {};
  ROWS().forEach(d => {byName[d.name] = d;});
}
const unitCost = d => {
  const y = Y(), b = y.disc(d), q = y.qty(d);
  return (has(b) && has(q) && n(q) > 0) ? n(b) / n(q) * 1000 : null;
};

/* ---------- plain-English scenario parser ---------- */
let LEDGER = [];
const norm = s => String(s || '').toLowerCase().replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();
const keyify = s => norm(s).replace(/ /g, '');
const ALIAS = {
  thaad: 'Terminal High Altitude Area Defense', ltamds: 'PATRIOT Lower-Tier Air and Missile Defense Sensor',
  pac3: 'PAC-3 Missile Segment Enhancement', gmd: 'Ground-based Midcourse Defense',
  aegis: 'Aegis Sea-Based Weapon System', shorad: 'Maneuver SHORAD',
  ifpc: 'Indirect Fire Protection Capability', virginia: 'Virginia-class Submarine',
  ssn: 'Virginia-class Submarine', columbia: 'Columbia-class Ballistic Missile Submarine',
  ssbn: 'Columbia-class Ballistic Missile Submarine', ddg: 'Arleigh Burke-class Destroyer',
  burke: 'Arleigh Burke-class Destroyer', destroyer: 'Arleigh Burke-class Destroyer',
  cvn: 'Gerald R. Ford-class Nuclear Aircraft Carrier', ford: 'Gerald R. Ford-class Nuclear Aircraft Carrier',
  carrier: 'Gerald R. Ford-class Nuclear Aircraft Carrier', lpd: 'San Antonio-class Amphibious Ship',
  lha: 'America-class Amphibious Assault Ship', lsm: 'Medium Landing Ship',
  frigate: 'Constellation-class Frigate / FF(X)', ffx: 'Constellation-class Frigate / FF(X)',
  usv: 'Unmanned Surface Vessels', uuv: 'Unmanned Undersea Vehicles',
  sentinel: 'LGM-35A Sentinel', icbm: 'LGM-35A Sentinel',
  trident: 'Trident II Ballistic Missile Modifications', lrso: 'Long-Range Standoff Weapon',
  jassm: 'Joint Air-to-Surface Standoff Missile', lrasm: 'Long-Range Anti-Ship Missile',
  amraam: 'Advanced Medium Range Air-to-Air Missile', sm6: 'Standard Missile-6',
  essm: 'Sea Sparrow', ram: 'Rolling Airframe Missile', nsm: 'Naval Strike Missile',
  prsm: 'Precision Strike Missile', gmlrs: 'Guided Multiple Launch Rocket System',
  typhon: 'Strategic Mid-Range Fires System / Typhon', tomahawk: 'Tactical Tomahawk Cruise Missile',
  javelin: 'Javelin', jagm: 'Joint Air-to-Ground Missile', sdb: 'Small Diameter Bomb II',
  ngad: 'F-47 Next Generation Air Dominance', f47: 'F-47 Next Generation Air Dominance',
  faxx: 'F/A-XX Next Generation Fighter', cca: 'Collaborative Combat Aircraft',
  wedgetail: 'E-7A Wedgetail', hawkeye: 'E-2D Advanced Hawkeye', raider: 'B-21 Raider',
  pegasus: 'KC-46A Pegasus', hercules: 'C-130J Hercules', osprey: 'V-22 Osprey',
  apache: 'AH-64 Apache', chinook: 'CH-47 Chinook', blackhawk: 'UH-60 Black Hawk',
  flraa: 'Future Long-Range Assault Aircraft', stingray: 'MQ-25 Stingray',
  stryker: 'Stryker Family of Armored Vehicles', ampv: 'Armored Multi-Purpose Vehicle',
  acv: 'Amphibious Combat Vehicle', jltv: 'Joint Light Tactical Vehicle',
  xm30: 'XM30 Combat Vehicle', ngsw: 'Next Generation Squad Weapon',
  hacm: 'Hypersonic Attack Cruise Missile', lrhw: 'Long-Range Hypersonic Weapon',
  darkeagle: 'Long-Range Hypersonic Weapon', cps: 'Conventional Prompt Strike',
  pnt: 'Position, Navigation, and Timing', nssl: 'Launch Enterprise'
};
const CATALIAS = {munitions: 'Missiles & Munitions', missiles: 'Missiles & Munitions',
  shipbuilding: 'Shipbuilding & Maritime', ships: 'Shipbuilding & Maritime',
  maritime: 'Shipbuilding & Maritime', aircraft: 'Aircraft', air: 'Aircraft',
  planes: 'Aircraft', aviation: 'Aircraft', space: 'Space', ground: 'Ground',
  vehicles: 'Ground', hypersonics: 'Hypersonic', hypersonic: 'Hypersonic',
  'missile defense': 'Missile Defense', c4i: 'C4I', networks: 'C4I', communications: 'C4I'};
const SVCALIAS = {navy: 'Navy', army: 'Army', 'air force': 'Air Force', airforce: 'Air Force',
  marines: 'Marine Corps', 'marine corps': 'Marine Corps', usmc: 'Marine Corps',
  'space force': 'Space Force', mda: 'MDA', 'missile defense agency': 'MDA',
  socom: 'SOCOM', joint: 'Multi-service', 'multi service': 'Multi-service'};

function idsWhere(field, value) {
  return ROWS().filter(d => d[field] === value).map(d => d.id);
}
function scoreName(a, q) {
  if (a === q) return 1000;
  if (a.indexOf(q) === 0) return 700 + q.length * 2;
  if (a.indexOf(q) >= 0) return 500 + q.length * 2;
  const at = a.split(' '), bt = q.split(' ');
  let hits = 0;
  bt.forEach(tk => {if (tk.length > 2 && at.some(x => x.indexOf(tk) === 0)) hits++;});
  return hits ? 120 * hits + Math.round(60 * hits / bt.length) : 0;
}
function matchTarget(raw) {
  let q = norm(raw).replace(/^(the|all|every|our|a|an)\s+/, '')
    .replace(/\s+(program|programs|programme|programmes|line|lines|account|accounts|budget|funding|spending|procurement)$/, '').trim();
  if (!q) return null;
  if (/^(everything|all|all systems|the portfolio|all lines|the whole portfolio|total)$/.test(q))
    return {scope: 'all', ids: ROWS().map(d => d.id), label: 'every selected line'};
  const k = keyify(q), k2 = k.replace(/s$/, ''), q2 = q.replace(/s$/, '');
  if (ALIAS[k] && byName[ALIAS[k]]) return {scope: 'system', ids: [byName[ALIAS[k]].id], label: ALIAS[k]};
  if (ALIAS[k2] && byName[ALIAS[k2]]) return {scope: 'system', ids: [byName[ALIAS[k2]].id], label: ALIAS[k2]};
  for (let i = 0; i < CATS.length; i++)
    if (norm(CATS[i]) === q) return {scope: 'category', ids: idsWhere('cat', CATS[i]), label: CATS[i]};
  const ca = CATALIAS[q] || CATALIAS[q2];
  if (ca) return {scope: 'category', ids: idsWhere('cat', ca), label: ca};
  for (let i = 0; i < SVCS.length; i++)
    if (norm(SVCS[i]) === q) return {scope: 'service', ids: idsWhere('svc', SVCS[i]), label: SVCS[i] + ' lines'};
  const sa = SVCALIAS[q] || SVCALIAS[q2];
  if (sa) return {scope: 'service', ids: idsWhere('svc', sa), label: sa + ' lines'};
  let scored = ROWS().map(d => ({d: d, s: scoreName(norm(d.name), q)})).filter(x => x.s > 0).sort((a, b) => b.s - a.s);
  if (!scored.length && q2 !== q)
    scored = ROWS().map(d => ({d: d, s: scoreName(norm(d.name), q2)})).filter(x => x.s > 0).sort((a, b) => b.s - a.s);
  if (!scored.length) return null;
  if (scored.length > 1 && scored[0].s < 500 && scored[1].s >= scored[0].s * 0.9)
    return {scope: 'ambiguous', options: scored.slice(0, 4).map(x => x.d.name)};
  return {scope: 'system', ids: [scored[0].d.id], label: scored[0].d.name};
}
function parseAmount(str) {
  const s = String(str || '').toLowerCase().replace(/[$,]/g, ' ').replace(/\s+/g, ' ').trim();
  let m = s.match(/(-?\d+(?:\.\d+)?)\s*(?:%|percent|pct)/);
  if (m) return {pct: parseFloat(m[1]) / 100};
  m = s.match(/(-?\d+(?:\.\d+)?)\s*(trillion|billion|million|thousand|tn|bn|mn|[btmk])?\b/);
  if (!m) return null;
  const mult = {trillion: 1000, tn: 1000, t: 1000, billion: 1, bn: 1, b: 1,
                million: 0.001, mn: 0.001, m: 0.001, thousand: 1e-6, k: 1e-6};
  const u = m[2] || '';
  return {usd: parseFloat(m[1]) * (u && mult[u] !== undefined ? mult[u] : 1), unit: u};
}
function parseLine(raw) {
  const line = String(raw || '').trim();
  if (!line || /^(#|\/\/)/.test(line)) return null;
  const out = {raw: line, ok: false, msg: '', purpose: '', dollars: 0};
  let s = line.replace(/[.]+$/, '');
  let offText = null;
  let m = s.match(/\b(?:offset(?:\s+by|\s+from|\s+with)?|paid\s+for\s+by|pay\s+for\s+(?:it|this)\s+(?:by|from|with)|funded\s+by|taken\s+from|take\s+it\s+from)\s+(?:cutting\s+|reducing\s+|trimming\s+|cuts\s+to\s+)?(.+)$/i);
  if (m) {offText = m[1]; s = s.slice(0, m.index).replace(/[,;]\s*$/, '').trim();}
  if (/\b(to|from|into|onto|out of)\b/i.test(s)) {
    m = s.match(/\bfor\s+(.+?)$/i);
    if (m) {out.purpose = m[1].replace(/\s+purposes?$/i, '').replace(/[,;]$/, '').trim(); s = s.slice(0, m.index).trim();}
  }
  let mode = null, amtStr = null, tgtStr = null, mm;
  if ((mm = s.match(/^(?:please\s+)?(?:add|increase|boost|plus|raise|grow|give|fund|append|put|move)\s+(.+?)\s+(?:to|into|onto|toward|towards|for)\s+(.+)$/i))) {
    mode = 'add'; amtStr = mm[1]; tgtStr = mm[2];
  } else if ((mm = s.match(/^(?:please\s+)?(?:cut|reduce|trim|slash|remove|subtract|take|strip|drop|defund)\s+(.+?)\s+(?:from|out of|off of|off|in)\s+(.+)$/i))) {
    mode = 'cut'; amtStr = mm[1]; tgtStr = mm[2];
  } else if ((mm = s.match(/^(?:please\s+)?(increase|raise|grow|boost|decrease|reduce|lower|cut|trim|change|adjust)\s+(.+?)\s+by\s+(.+)$/i))) {
    mode = /^(decrease|reduce|lower|cut|trim)$/i.test(mm[1]) ? 'cut' : 'add'; tgtStr = mm[2]; amtStr = mm[3];
  } else if ((mm = s.match(/^(?:please\s+)?(?:set|make|hold|peg|fix)\s+(.+?)\s+(?:to|at)\s+(.+)$/i))) {
    mode = 'set'; tgtStr = mm[1]; amtStr = mm[2];
  } else if ((mm = s.match(/^(?:please\s+)?(?:zero|eliminate|cancel|terminate|kill|end|defund)\s+(?:out\s+)?(.+)$/i))) {
    mode = 'zero'; tgtStr = mm[1];
  } else {
    const a = s.match(/(-?\$?\s*\d[\d,]*\.?\d*\s*(?:%|percent|pct|trillion|billion|million|thousand|bn|mn|[btmk])?)/i);
    if (a) {
      amtStr = a[1];
      tgtStr = (s.slice(0, a.index) + ' ' + s.slice(a.index + a[1].length))
        .replace(/\b(add|increase|boost|cut|reduce|trim|to|from|for|by|of|in|on|the|a|an|please)\b/gi, ' ').trim();
      mode = /\b(cut|reduce|trim|slash|remove|defund|drop|lower|decrease)\b/i.test(s) ? 'cut' : 'add';
    }
  }
  if (!mode) {out.msg = 'I could not tell what action you want. Try "add 1 billion to ..." or "cut 10% from ...".'; return out;}
  const amt = mode === 'zero' ? {usd: 0} : parseAmount(amtStr || '');
  if (mode !== 'zero' && !amt) {out.msg = 'I could not find an amount in that line.'; return out;}
  const t = matchTarget(tgtStr || '');
  if (!t) {out.msg = 'I could not tell which line you mean.'; return out;}
  if (t.scope === 'ambiguous') {
    out.msg = 'That could be ' + t.options.join(', ') + '. Add a word or two.';
    out.options = t.options; return out;
  }
  out.ok = true; out.ids = t.ids; out.label = t.label; out.scopeKind = t.scope;
  if (mode === 'zero') {out.mode = 'zero';}
  else if (amt.pct !== undefined) {out.mode = 'pct'; out.pct = (mode === 'cut' ? -1 : 1) * Math.abs(amt.pct);}
  else if (mode === 'set') {out.mode = 'set'; out.usd = Math.abs(amt.usd);}
  else {out.mode = 'usd'; out.usd = (mode === 'cut' ? -1 : 1) * Math.abs(amt.usd); out.assumedUnit = !amt.unit;}
  if (offText) {
    const o = matchTarget(offText);
    if (o && o.ids) {out.offsetIds = o.ids; out.offsetLabel = o.label;}
    else out.offsetMsg = 'I could not find the offset target.';
  }
  out.reading = (out.mode === 'zero' ? 'Zero out ' + out.label
    : out.mode === 'set' ? 'Set ' + out.label + ' to ' + money(out.usd)
    : out.mode === 'pct' ? (out.pct >= 0 ? 'Add ' : 'Cut ') + P(Math.abs(out.pct)) + (out.pct >= 0 ? ' to ' : ' from ') + out.label
    : (out.usd >= 0 ? 'Add ' : 'Cut ') + money(Math.abs(out.usd)) + (out.usd >= 0 ? ' to ' : ' from ') + out.label);
  return out;
}
const parseScript = text => String(text || '').split(/\n+/).map(parseLine).filter(x => x);

function ledgerEffects(base0) {
  const pct = {}, usd = {}, entries = parseScript(S.text || '');
  entries.forEach(e => {
    if (!e.ok) return;
    const ids = e.ids || [];
    if (!ids.length) {e.ok = false; e.msg = 'No lines matched.'; return;}
    const pool = ids.reduce((t, id) => t + (base0[id] || 0), 0);
    let dollars = 0;
    if (e.mode === 'pct') {
      ids.forEach(id => {pct[id] = (pct[id] || 0) + e.pct;});
      dollars = pool * e.pct;
    } else if (e.mode === 'zero') {
      ids.forEach(id => {usd[id] = (usd[id] || 0) - (base0[id] || 0);});
      dollars = -pool;
    } else if (e.mode === 'set') {
      dollars = e.usd - pool;
      ids.forEach(id => {usd[id] = (usd[id] || 0) + (pool ? dollars * (base0[id] || 0) / pool : dollars / ids.length);});
    } else {
      dollars = e.usd;
      ids.forEach(id => {usd[id] = (usd[id] || 0) + (pool ? e.usd * (base0[id] || 0) / pool : e.usd / ids.length);});
    }
    e.dollars = dollars;
    if (e.offsetIds && e.offsetIds.length) {
      const opool = e.offsetIds.reduce((t, id) => t + (base0[id] || 0), 0);
      e.offsetIds.forEach(id => {
        usd[id] = (usd[id] || 0) - (opool ? dollars * (base0[id] || 0) / opool : dollars / e.offsetIds.length);
      });
    }
  });
  LEDGER = entries;
  return {pct: pct, usd: usd, entries: entries};
}
"""

TEMPLATE_JS2 = r"""
/* ---------- state ---------- */
const PAY = DATA.pay;
const GRADES = PAY.grades.map(g => g.g);
const REPYOS = PAY.repyos;
const A = {};
PAY.assumptions.forEach(a => {A[a.k] = a.v;});

const DEF = () => ({
  year: 2027, tab: 'overview',
  baseline: {2026: 'app', 2027: 'disc'},
  global: 0, sys: {}, offset: 'none', offsetSys: '', text: '',
  claim: {sys: '', base: '', type: 'pct', val: 0.10, unit: 1e9, off: 'none',
          offSys: '', ref: '', own: false, ownLab: ''},
  pay: {year: 2027, ata: 0, month: 1, rpaOn: 1, svc: 'All services',
        grade: 'E-4', gradePct: 0.10},
  proj: {infl: DATA.ctx.proj.infl, real: DATA.ctx.proj.real, repl: DATA.ctx.proj.repl},
  dview: {q: '', cat: 'all', svc: 'all', sort: 'disc', dir: -1},
  cview: {q: '', cat: 'all', pres: 'all', sort: 'dTot', dir: -1},
  sview: {sort: 'delta', dir: -1}
});
let S = DEF();

const enc = o => btoa(unescape(encodeURIComponent(JSON.stringify(o))));
const dec = s => JSON.parse(decodeURIComponent(escape(atob(s))));
function persist() {
  let s;
  try {s = enc(S);} catch (e) {return;}
  /* Some browsers refuse to rewrite the address of a file:// page. That must not
     stop the local save, so the two are attempted separately. */
  try {history.replaceState(null, '', '#' + s);} catch (e) {}
  try {localStorage.setItem('fy2627lab', s);} catch (e) {}
}
function restore() {
  try {
    const raw = (location.hash || '').slice(1) || localStorage.getItem('fy2627lab');
    if (!raw) return;
    const o = dec(raw), d = DEF();
    S = Object.assign(d, o);
    ['claim', 'pay', 'proj', 'dview', 'cview', 'sview'].forEach(k => {
      S[k] = Object.assign(DEF()[k], o[k] || {});
    });
    S.baseline = Object.assign(DEF().baseline, o.baseline || {});
    YEAR = (S.year === 2026 || S.year === 2027) ? S.year : 2027;
  } catch (e) {}
}

/* ---------- scenario engine ---------- */
function compute() {
  const bl = BASES();
  const key = bl[S.baseline[YEAR]] ? S.baseline[YEAR] : Y().defBase;
  const bf = bl[key].f;
  const rows0 = ROWS();
  const base0 = {};
  rows0.forEach(d => {base0[d.id] = bf(d);});
  const L = ledgerEffects(base0);
  const rows = rows0.map(d => {
    const b = base0[d.id], a = S.sys[d.id] || {p: 0, d: 0};
    const lp = L.pct[d.id] || 0, ld = L.usd[d.id] || 0;
    const touched = n(a.p) !== 0 || n(a.d) !== 0 || S.global !== 0 || lp !== 0 || ld !== 0;
    const scen = b * (1 + S.global) * (1 + n(a.p) + lp) + n(a.d) + ld;
    return {d: d, b: b, scen: scen, touched: touched, p: n(a.p), amt: n(a.d),
            lp: lp, ld: ld, off: 0};
  });
  const net = sum(rows, r => r.scen - r.b);
  if (S.offset === 'all') {
    const pool = sum(rows.filter(r => !r.touched), r => r.b);
    if (pool > 0) rows.forEach(r => {if (!r.touched) r.off = -net * r.b / pool;});
  } else if (S.offset === 'cat') {
    CATS.forEach(c => {
      const rs = rows.filter(r => r.d.cat === c);
      const cnet = sum(rs, r => r.scen - r.b);
      const pool = sum(rs.filter(r => !r.touched), r => r.b);
      if (pool > 0 && Math.abs(cnet) > 1e-12) rs.forEach(r => {if (!r.touched) r.off = -cnet * r.b / pool;});
    });
  } else if (S.offset === 'sys' && S.offsetSys) {
    const t = rows.find(r => r.d.name === S.offsetSys);
    if (t) t.off = -net;
  }
  rows.forEach(r => {r.final = r.scen + r.off; r.delta = r.final - r.b;});
  return {rows: rows, baseKey: key, baseLabel: bl[key].label};
}

/* ---------- pay engine ----------
   Basic pay comes from the 2026 table the workbook carries. The burden rates,
   the caps and the composite comparison switch with the fiscal year, exactly as
   the Pay_Module assumptions switch. */
const PAYYEAR = () => (S.pay.year === 2026 ? 2026 : 2027);
function payRates() {
  const y = PAYYEAR();
  const m1 = {};
  PAY.m1.forEach(r => {m1[r.item] = (y === 2026 ? r.fy26 : r.fy27);});
  return {
    rpa: n(m1['Implied retired pay accrual rate']),
    tspMatch: n(m1['Implied TSP matching rate']),
    tspAuto: n(A['TSP automatic agency contribution']),
    med: n(A['Employer Medicare (HI) rate']),
    oasdi: n(A['Employer Social Security (OASDI) rate']),
    wage: n(A['OASDI taxable maximum']),
    capV: n(A['Monthly cap, O-6 and below']),
    capII: n(A['Monthly cap, O-7 to O-10'])
  };
}
const payCap = (g, R) => (g.charAt(0) === 'O' && +g.slice(2) >= 7) ? R.capII : R.capV;
function payRate(g) {
  const row = PAY.table[g];
  if (!row) return 0;
  const i = PAY.steps.indexOf(REPYOS[g]);
  return i < 0 ? 0 : n(row[i]);
}
function strengthFor(g, svc) {
  if (!svc || svc === 'All services') return n(PAY.force[g]);
  return n((PAY.forcesvc[g] || {})[svc]);
}
function payCompute(svcOpt) {
  const p = S.pay, R = payRates();
  const months = p.month >= 10 ? 22 - p.month : 10 - p.month;
  const svc = svcOpt || p.svc || 'All services';
  const tsp = R.tspAuto + R.tspMatch;
  const rows = GRADES.map(g => {
    const strength = strengthFor(g, svc), pay = payRate(g), cap = payCap(g, R);
    const extra = (g === p.grade) ? n(p.gradePct) : 0;
    const tot = n(p.ata) + extra, cur = Math.min(pay, cap);
    const nw = pay === 0 ? 0 : Math.min(cap, pay * (1 + tot));
    const inc = pay === 0 ? 0 : nw - cur, annual = inc * 12;
    const payroll = strength * cur * 12 / 1e9;
    const basic = strength * annual / 1e9;
    const rpa = p.rpaOn ? basic * R.rpa : 0;
    const fica = strength * (R.med * annual
      + R.oasdi * (Math.min(nw * 12, R.wage) - Math.min(cur * 12, R.wage))) / 1e9;
    const tspC = basic * tsp;
    const total = basic + rpa + fica + tspC;
    const one = payroll * 0.01 * (1 + (p.rpaOn ? R.rpa : 0) + tsp + R.med)
      + strength * R.oasdi * (Math.min(cur * 12 * 1.01, R.wage) - Math.min(cur * 12, R.wage)) / 1e9;
    return {g: g, strength: strength, pay: pay, cap: cap, tot: tot, nw: nw, inc: inc,
            annual: annual, payroll: payroll, basic: basic, rpa: rpa, fica: fica,
            tsp: tspC, total: total, first: total * months / 12, one: one};
  });
  const T = k => sum(rows, r => r[k]);
  const affected = sum(rows.filter(r => Math.abs(r.inc) > 0.0001), r => r.strength);
  return {rows: rows, svc: svc, months: months, rates: R, total: T('total'),
          first: T('first'), payroll: T('payroll'), one: T('one'), basic: T('basic'),
          affected: affected, avg: affected ? T('basic') * 1e9 / affected : 0};
}

function projCompute() {
  const g = (1 + n(S.proj.infl)) * (1 + n(S.proj.real));
  return DATA.ctx.projRows.map(r => {
    const gaps = [1, 2, 3, 4, 5].map(t => (1 - n(S.proj.repl)) * n(r.spend) * Math.pow(g, t));
    return {c: r.cat, app: n(r.app), spend: n(r.spend), tot: n(r.tot),
            dep: n(r.dep), growth: r.app ? n(r.tot) * g / n(r.app) - 1 : null,
            gaps: gaps, cum: gaps.reduce((a, b) => a + b, 0)};
  });
}

/* The typed amount is held in whatever unit the reader picked. Everything inside
   the model is billions of dollars, so it is converted once, here. */
const claimBillions = () => n(S.claim.val) * n(S.claim.unit || 1e9) / 1e9;

function claimCompute() {
  const c = S.claim;
  const bl = BASES(), key = bl[c.base] ? c.base : Y().defBase, bf = bl[key].f;
  const allTotal = sum(ROWS(), bf);
  const pay = payCompute();

  /* A claim of the reader's own is not a line in either report. It has no
     baseline of its own, no quantity and no unit cost, so those are left blank
     rather than invented. Everything that can still be said about it is said:
     its size against the portfolio, what it would take out of an offset pool,
     what it buys in units of a line that does have a published quantity, and
     what it is worth as a pay raise. */
  if (c.own) {
    const chg0 = claimBillions();
    const ref0 = byName[c.ref], refUC0 = ref0 ? unitCost(ref0) : null;
    let pool0 = 0, poolLabel0 = '';
    if (c.off === 'sys') {pool0 = byName[c.offSys] ? bf(byName[c.offSys]) : 0; poolLabel0 = c.offSys;}
    else if (c.off === 'cat') {pool0 = allTotal; poolLabel0 = 'the selected lines as a whole';}
    else if (c.off === 'all') {pool0 = allTotal; poolLabel0 = 'every selected line';}
    return {own: true, d: null, label: c.ownLab || 'a claim of your own',
            b: null, chg: chg0, nw: null, baseLabel: bl[key].label,
            catTotal: null, allTotal: allTotal, uc: null, ref: ref0, refUC: refUC0,
            pool: pool0, poolLabel: poolLabel0, pctCut: pool0 ? -chg0 / pool0 : null,
            units: null, refUnits: refUC0 ? chg0 * 1000 / refUC0 : null,
            raise: pay.one > 0 ? chg0 / pay.one * 0.01 : null};
  }

  const d = byName[c.sys];
  if (!d) return null;
  const b = bf(d);
  const chg = c.type === 'pct' ? b * n(c.val) : claimBillions();
  const catTotal = sum(ROWS().filter(x => x.cat === d.cat), bf);
  const uc = unitCost(d), ref = byName[c.ref], refUC = ref ? unitCost(ref) : null;
  let pool = 0, poolLabel = '';
  if (c.off === 'sys') {pool = byName[c.offSys] ? bf(byName[c.offSys]) : 0; poolLabel = c.offSys;}
  else if (c.off === 'cat') {pool = catTotal - b; poolLabel = 'the rest of ' + d.cat;}
  else if (c.off === 'all') {pool = allTotal - b; poolLabel = 'every other selected line';}
  return {own: false, d: d, label: d.name, b: b, chg: chg, nw: b + chg,
          baseLabel: bl[key].label, catTotal: catTotal, allTotal: allTotal,
          uc: uc, ref: ref, refUC: refUC,
          pool: pool, poolLabel: poolLabel, pctCut: pool ? -chg / pool : null,
          units: uc ? chg * 1000 / uc : null, refUnits: refUC ? chg * 1000 / refUC : null,
          raise: pay.one > 0 ? chg / pay.one * 0.01 : null};
}

/* A small amount in a large table is easier to read in its own units than in
   billions, so both are offered. */
function amountWords(bn) {
  const d = Math.abs(bn) * 1e9;
  if (d === 0) return '$0';
  if (d < 1000) return '$' + N(d, d < 10 ? 2 : 0);
  if (d < 1e6) return '$' + N(d / 1e3, 1) + ' thousand';
  if (d < 1e9) return '$' + N(d / 1e6, d < 1e7 ? 2 : 1) + ' million';
  if (d < 1e12) return '$' + N(d / 1e9, 2) + ' billion';
  return '$' + N(d / 1e12, 2) + ' trillion';
}

/* ---------- charts, hand written ---------- */
function nice(x) {
  if (!(x > 0)) return 1;
  const e = Math.pow(10, Math.floor(Math.log10(x))), f = x / e;
  return (f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10) * e;
}
const PAL = ['#a8506e', '#d98ba3', '#c9b6d9', '#4a7f92', '#3b3354', '#f2b8c6', '#9a6b1f', '#2f7d5b'];
function vChart(cats, series, o) {
  o = o || {};
  const W = o.w || 780, H = o.h || 300, L = o.L || 56, R = 14, T = 14;
  const Bm = series.length > 1 ? 58 : 42;
  let mx = 0, mn = 0;
  cats.forEach((c, i) => {
    if (o.stack) {
      let p = 0, q = 0;
      series.forEach(s => {const v = +s.data[i] || 0; if (v >= 0) p += v; else q += v;});
      mx = Math.max(mx, p); mn = Math.min(mn, q);
    } else series.forEach(s => {const v = +s.data[i] || 0; mx = Math.max(mx, v); mn = Math.min(mn, v);});
  });
  mx = nice(mx); mn = mn < 0 ? -nice(-mn) : 0;
  const H0 = H - T - Bm, y = v => T + H0 * (mx - v) / ((mx - mn) || 1);
  let g = '';
  for (let i = 0; i <= 4; i++) {
    const v = mn + (mx - mn) * i / 4, yy = y(v);
    g += '<line class="grid" x1="' + L + '" x2="' + (W - R) + '" y1="' + yy + '" y2="' + yy + '"/>'
      + '<text class="ax" x="' + (L - 6) + '" y="' + (yy + 3) + '" text-anchor="end">'
      + (o.fmt ? o.fmt(v) : v.toFixed(0)) + '</text>';
  }
  if (mn < 0) g += '<line class="zero" x1="' + L + '" x2="' + (W - R) + '" y1="' + y(0) + '" y2="' + y(0) + '"/>';
  const bw = (W - L - R) / cats.length, inner = bw * 0.74;
  const sw = o.stack ? inner : inner / series.length;
  cats.forEach((c, i) => {
    const x0 = L + bw * i + (bw - inner) / 2;
    if (o.stack) {
      let acc = 0;
      series.forEach(s => {
        const v = +s.data[i] || 0, y1 = y(acc + v), h = Math.abs(y(acc) - y(acc + v));
        acc += v;
        g += '<rect x="' + x0 + '" y="' + y1 + '" width="' + inner + '" height="' + Math.max(h, 0)
          + '" fill="' + s.color + '"><title>' + esc(c) + ' ' + esc(s.name) + ': '
          + (o.fmt ? o.fmt(v) : v) + '</title></rect>';
      });
    } else {
      series.forEach((s, j) => {
        const v = +s.data[i] || 0, y1 = Math.min(y(v), y(0)), h = Math.abs(y(v) - y(0));
        g += '<rect x="' + (x0 + sw * j) + '" y="' + y1 + '" width="' + Math.max(sw - 2, 1)
          + '" height="' + Math.max(h, 1) + '" fill="' + s.color + '"><title>' + esc(c) + ' '
          + esc(s.name) + ': ' + (o.fmt ? o.fmt(v) : v) + '</title></rect>';
      });
    }
    const lab = String(c).length > 13 ? String(c).slice(0, 12) + '.' : c;
    g += '<text class="ax" x="' + (L + bw * i + bw / 2) + '" y="' + (H - Bm + 15)
      + '" text-anchor="middle">' + esc(lab) + '<title>' + esc(c) + '</title></text>';
  });
  if (series.length > 1) {
    const step = Math.min(170, (W - L) / series.length);
    g += series.map((s, i) => '<g transform="translate(' + (L + i * step) + ',' + (H - 12) + ')">'
      + '<rect width="9" height="9" y="-8" fill="' + s.color + '"/>'
      + '<text class="ax" x="13">' + esc(s.name) + '</text></g>').join('');
  }
  return '<svg class="svg anim" viewBox="0 0 ' + W + ' ' + H + '" role="img">' + g + '</svg>';
}
function hChart(items, o) {
  o = o || {};
  const W = o.w || 780, rowH = o.rowH || 21, L = o.L || 250, R = 84, T = 8;
  const H = T + items.length * rowH + 8;
  const mx = nice(Math.max.apply(null, items.map(i => Math.abs(n(i.value))).concat([0.0001])));
  const neg = items.some(i => n(i.value) < 0);
  const span = W - L - R, zero = neg ? L + span / 2 : L, sc = (neg ? span / 2 : span) / mx;
  let g = '';
  items.forEach((it, i) => {
    const v = n(it.value), y = T + i * rowH, w = Math.max(Math.abs(v) * sc, 1);
    const x = v < 0 ? zero - w : zero;
    const lab = String(it.label).length > 34 ? String(it.label).slice(0, 33) + '.' : it.label;
    g += '<text class="ax" x="' + (L - 8) + '" y="' + (y + 13) + '" text-anchor="end">'
      + esc(lab) + '<title>' + esc(it.label) + '</title></text>'
      + '<rect x="' + x + '" y="' + (y + 4) + '" width="' + w + '" height="' + (rowH - 9)
      + '" fill="' + (it.color || (v < 0 ? '#b0304a' : '#d98ba3')) + '"/>'
      + '<text class="ax" x="' + ((v < 0 ? zero : zero + w) + 6) + '" y="' + (y + 13) + '">'
      + (o.fmt ? o.fmt(v) : v) + '</text>';
  });
  if (neg) g += '<line class="zero" x1="' + zero + '" x2="' + zero + '" y1="' + T + '" y2="' + (H - 8) + '"/>';
  return '<svg class="svg anim h" viewBox="0 0 ' + W + ' ' + H + '" role="img">' + g + '</svg>';
}
function donut(parts, o) {
  o = o || {};
  const W = o.w || 300, H = o.h || 300, cx = W / 2, cy = H / 2;
  const r = Math.min(W, H) / 2 - 8, ri = r * 0.58;
  const tot = sum(parts, p => Math.max(n(p.value), 0)) || 1;
  let a0 = -Math.PI / 2, g = '';
  parts.forEach((p, i) => {
    const frac = Math.max(n(p.value), 0) / tot, a1 = a0 + frac * Math.PI * 2;
    const large = (a1 - a0) > Math.PI ? 1 : 0;
    const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
    const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
    const xi1 = cx + ri * Math.cos(a1), yi1 = cy + ri * Math.sin(a1);
    const xi0 = cx + ri * Math.cos(a0), yi0 = cy + ri * Math.sin(a0);
    if (frac > 0.00001) {
      g += '<path d="M' + x0 + ' ' + y0 + ' A' + r + ' ' + r + ' 0 ' + large + ' 1 ' + x1 + ' ' + y1
        + ' L' + xi1 + ' ' + yi1 + ' A' + ri + ' ' + ri + ' 0 ' + large + ' 0 ' + xi0 + ' ' + yi0
        + ' Z" fill="' + (p.color || PAL[i % PAL.length]) + '"><title>' + esc(p.label) + ': '
        + (o.fmt ? o.fmt(n(p.value)) : n(p.value)) + ' (' + P(frac) + ')</title></path>';
    }
    a0 = a1;
  });
  if (o.center) {
    g += '<text class="axb" x="' + cx + '" y="' + (cy - 2) + '" text-anchor="middle" font-size="19">'
      + esc(o.center) + '</text>';
    if (o.sub) g += '<text class="ax" x="' + cx + '" y="' + (cy + 16) + '" text-anchor="middle">'
      + esc(o.sub) + '</text>';
  }
  return '<svg class="svg anim" viewBox="0 0 ' + W + ' ' + H + '" role="img">' + g + '</svg>';
}
"""

TEMPLATE_JS3 = r"""
/* ---------- rendering ---------- */
function opts(list, sel) {
  return list.map(o => {
    const v = (typeof o === 'string') ? o : o.v, l = (typeof o === 'string') ? o : o.l;
    return '<option value="' + esc(v) + '"' + (v === sel ? ' selected' : '') + '>' + esc(l) + '</option>';
  }).join('');
}
const fig = (v, k, s, words) => '<div class="fig"><div class="v' + (words ? ' words' : '')
  + '">' + v + '</div><div class="k">' + k + '</div>'
  + (s ? '<div class="s">' + esc(s) + '</div>' : '') + '</div>';
function th(label, key, cls, sortable) {
  const v = S[TBL] || {};
  const on = sortable && v.sort === key;
  return '<th class="' + (cls || '') + (sortable ? ' sortable' : '') + '"'
    + (sortable ? ' data-sort="' + esc(key) + '"' : '') + '>' + esc(label)
    + (on ? '<span class="arrow">' + (v.dir < 0 ? 'v' : '^') + '</span>' : '') + '</th>';
}
let TBL = 'dview';

function sortRows(rows, view, get) {
  const k = view.sort, dir = view.dir;
  return rows.slice().sort((a, b) => {
    const va = get(a, k), vb = get(b, k);
    const na = (va === null || va === undefined || va === '');
    const nb = (vb === null || vb === undefined || vb === '');
    if (na && nb) return 0;
    if (na) return 1;
    if (nb) return -1;
    if (typeof va === 'string' || typeof vb === 'string')
      return dir * String(va).localeCompare(String(vb));
    return dir * (va - vb);
  });
}

/* ---------- overview ---------- */
function renderOverview() {
  const y = Y(), C = y.cats, rows = C.rows, tot = C.total;
  const isBoth = S.year === 'both';
  $('#heroKicker').textContent = isBoth ? 'Fiscal years 2026 and 2027' : y.label;
  $('#heroLead').innerHTML = isBoth
    ? 'Two Congressional Research Service reports, one spine. FY2026 is settled: an authorization and an appropriation are law. FY2027 is not: the request has three committee marks against it and nothing enacted. Where a line appears in only one of the two reports it stays blank here, because a programme CRS did not list is not a programme cut to zero.'
    : (YEAR === 2027
      ? 'The FY2027 request asks ' + B(n(DATA.cats27.total.total), 1) + ' across the '
        + N(tot.n) + ' lines inside the eight category tables, of which '
        + B(n(DATA.cats27.total.mand), 1) + ' is mandatory money carried over from reconciliation. '
        + 'Three committees have reported. Nothing has been enacted, so every enacted column is blank.'
      : 'The FY2026 request asked ' + B(n(DATA.cats26.total.total), 1) + ' across '
        + N(tot.n) + ' selected systems, ' + B(n(DATA.cats26.total.recon), 1)
        + ' of it mandatory. Both the authorization and the appropriation are now law, and the '
        + 'reconciliation spend plan adds ' + B(n(DATA.cats26.total.spend), 1) + ' on top.');

  const figs = [];
  if (isBoth) {
    const t26 = n(DATA.cats26.total.total), t27 = n(DATA.cats27.total.total);
    figs.push(fig(B(t26, 1), 'FY2026 total request, 97 selected systems', 'R48860'));
    figs.push(fig(B(t27, 1), 'FY2027 total request, 101 lines in the category tables', 'R49023'));
    figs.push(fig(Bs(t27 - t26, 1), 'Change across every line listed, either year', 'not like for like'));
    const br = DATA.compare.bridge.find(b => /matched lines$/.test(b.step) && /^Change/.test(b.step));
    const priced = DATA.compare.rows.filter(r => r.presence === 'Both years'
      && has(r.tot26) && has(r.tot27));
    figs.push(fig(br ? Bs(n(br.amt), 1) : 'n/a',
      'Change on the ' + N(priced.length) + ' lines priced in both reports', 'like for like'));
  } else if (YEAR === 2027) {
    const t = DATA.cats27.total;
    figs.push(fig(B(n(t.total), 1), 'Total request, the eight category tables', 'R49023 pp. 4 to 22'));
    figs.push(fig(B(n(t.disc), 1), 'Discretionary', 'CRS narrative ' + B(n(t.discN), 1)));
    figs.push(fig(B(n(t.mand), 1), 'Mandatory, carried from reconciliation', 'CRS narrative ' + B(n(t.mandN), 1)));
    figs.push(fig(P(n(t.coverage)), 'Selected share of the category-wide request', B(n(t.program), 0) + ' category wide'));
  } else {
    const t = DATA.cats26.total;
    figs.push(fig(B(n(t.total), 1), 'Total request, 97 selected systems', 'R48860 Tables 1 to 7'));
    figs.push(fig(B(n(t.app), 1), 'Enacted appropriation, P.L. 119-75', Bs(n(t.appDelta), 1) + ' against the base request'));
    figs.push(fig(B(n(t.spend), 1), 'Reconciliation spend plan, P.L. 119-21', 'Appendix Table A-1'));
    figs.push(fig(P(n(t.mandDep)), 'Mandatory dependence of enacted resources', B(n(t.enactTotal), 1) + ' enacted in total'));
  }
  $('#heroFigs').innerHTML = figs.join('');

  /* hero chart */
  if (isBoth) {
    $('#heroVisT').textContent = 'Total request by category, both years, billions of dollars';
    $('#heroVis').innerHTML = vChart(CATS, [
      {name: 'FY2026', color: '#c9b6d9', data: DATA.cats26.rows.map(r => n(r.total))},
      {name: 'FY2027', color: '#a8506e', data: DATA.cats27.rows.map(r => n(r.total))}
    ], {h: 320, fmt: v => v.toFixed(0)});
  } else {
    $('#heroVisT').textContent = y.label + ' request by category, billions of dollars';
    const dk = y.catDisc, mk = y.catMand;
    $('#heroVis').innerHTML = vChart(CATS, [
      {name: 'Discretionary', color: '#a8506e', data: rows.map(r => n(r[dk]))},
      {name: 'Mandatory', color: '#c9b6d9', data: rows.map(r => n(r[mk]))}
    ], {stack: true, h: 320, fmt: v => v.toFixed(0)});
  }

  /* category table */
  if (isBoth) {
    $('#ovTitle').textContent = 'The eight categories, both years side by side';
    $('#ovLead').textContent = 'FY2026 columns are as enacted where an enacted figure exists. '
      + 'FY2027 columns are the request and the three committee marks. The two years are not '
      + 'like for like and the crosswalk section says exactly why.';
    $('#ovChartT').textContent = 'Total request by category, FY2026 against FY2027';
    $('#ovChart').innerHTML = hChart(CATS.map((c, i) => ({
      label: c, value: n(DATA.cats27.rows[i].total) - n(DATA.cats26.rows[i].total)
    })), {fmt: v => Bs(v, 1), L: 200});
    let h = '<thead><tr><th>Category</th><th class="num">FY2026 total request</th>'
      + '<th class="num">FY2026 enacted total</th><th class="num">FY2027 total request</th>'
      + '<th class="num">Change in request</th><th class="num">FY2027 HAC</th></tr></thead><tbody>';
    CATS.forEach((c, i) => {
      const a = DATA.cats26.rows[i], b = DATA.cats27.rows[i];
      const dv = n(b.total) - n(a.total);
      h += '<tr><td class="name">' + esc(c) + '</td>'
        + '<td class="num">' + B(n(a.total), 2) + '</td>'
        + '<td class="num">' + B(n(a.enactTotal), 2) + '</td>'
        + '<td class="num">' + B(n(b.total), 2) + '</td>'
        + '<td class="num ' + cl(dv) + '">' + Bs(dv, 2) + '</td>'
        + '<td class="num">' + cellB(b.hac, v => B(n(v), 2)) + '</td></tr>';
    });
    const A26 = DATA.cats26.total, A27 = DATA.cats27.total;
    h += '</tbody><tfoot><tr><td>All categories</td>'
      + '<td class="num">' + B(n(A26.total), 2) + '</td><td class="num">' + B(n(A26.enactTotal), 2) + '</td>'
      + '<td class="num">' + B(n(A27.total), 2) + '</td>'
      + '<td class="num ' + cl(n(A27.total) - n(A26.total)) + '">' + Bs(n(A27.total) - n(A26.total), 2) + '</td>'
      + '<td class="num">' + B(n(A27.hac), 2) + '</td></tr></tfoot>';
    $('#ovTable').innerHTML = h;
    $('#ovNote').textContent = 'The FY2026 column counts 97 systems and the FY2027 column counts '
      + '101 lines, so the difference in the last column mixes a change in funding with a change '
      + 'in what CRS chose to tabulate.';
  } else if (YEAR === 2027) {
    $('#ovTitle').textContent = 'The eight categories CRS tabulates for FY2027';
    $('#ovLead').textContent = 'Computed from the line items, set beside the figure CRS states in '
      + 'its own narrative. Where the two disagree the gap is shown rather than reconciled.';
    $('#ovChartT').textContent = 'Committee marks against the discretionary request, billions of dollars';
    $('#ovChart').innerHTML = vChart(CATS, [
      {name: 'House', color: '#a8506e', data: rows.map(r => n(r.houseDelta))},
      {name: 'SASC', color: '#c9b6d9', data: rows.map(r => n(r.sascDelta))},
      {name: 'HAC', color: '#4a7f92', data: rows.map(r => n(r.hacDelta))}
    ], {h: 300, fmt: v => v.toFixed(1)});
    let h = '<thead><tr><th>Category</th><th class="num">Lines</th>'
      + '<th class="num">Discretionary, computed</th><th class="num">CRS narrative</th>'
      + '<th class="num">Gap</th><th class="num">Mandatory</th><th class="num">Total request</th>'
      + '<th class="num">House</th><th class="num">SASC</th><th class="num">HAC</th>'
      + '<th class="num">Selected share</th></tr></thead><tbody>';
    rows.forEach(r => {
      h += '<tr><td class="name">' + esc(r.cat) + '</td><td class="num">' + N(r.n) + '</td>'
        + '<td class="num">' + B(n(r.disc), 2) + '</td><td class="num">' + B(n(r.discN), 1) + '</td>'
        + '<td class="num ' + cl(n(r.discGap)) + '">' + Bs(n(r.discGap), 2) + '</td>'
        + '<td class="num">' + (n(r.mand) ? B(n(r.mand), 2) : '<span class="dim">' + B(0, 2) + '</span>') + '</td>'
        + '<td class="num">' + B(n(r.total), 2) + '</td>'
        + '<td class="num ' + cl(n(r.houseDelta)) + '">' + B(n(r.house), 2) + '</td>'
        + '<td class="num ' + cl(n(r.sascDelta)) + '">' + B(n(r.sasc), 2) + '</td>'
        + '<td class="num ' + cl(n(r.hacDelta)) + '">' + B(n(r.hac), 2) + '</td>'
        + '<td class="num">' + P(n(r.coverage)) + '</td></tr>';
    });
    h += '</tbody><tfoot><tr><td>All eight categories</td><td class="num">' + N(tot.n) + '</td>'
      + '<td class="num">' + B(n(tot.disc), 2) + '</td><td class="num">' + B(n(tot.discN), 1) + '</td>'
      + '<td class="num ' + cl(n(tot.discGap)) + '">' + Bs(n(tot.discGap), 2) + '</td>'
      + '<td class="num">' + B(n(tot.mand), 2) + '</td><td class="num">' + B(n(tot.total), 2) + '</td>'
      + '<td class="num">' + B(n(tot.house), 2) + '</td><td class="num">' + B(n(tot.sasc), 2) + '</td>'
      + '<td class="num">' + B(n(tot.hac), 2) + '</td><td class="num">' + P(n(tot.coverage)) + '</td>'
      + '</tr></tfoot>';
    $('#ovTable').innerHTML = h;
    $('#ovNote').textContent = 'CRS rounds each line to the hundredth of a billion and each '
      + 'narrative total to the tenth, so a gap of a few hundredths is rounding. Nothing here is '
      + 'adjusted to close it. The three lines of Appendix Table A-1 sit outside these totals; '
      + 'they carry ' + B(n(DATA.cats27.a1total.mand), 1) + ' of mandatory money on their own.';
  } else {
    $('#ovTitle').textContent = 'The eight categories CRS tabulates for FY2026';
    $('#ovLead').textContent = 'Request, authorization, appropriation and the reconciliation spend '
      + 'plan, by category. The last two columns are law.';
    $('#ovChartT').textContent = 'Appropriation against the base request, billions of dollars';
    $('#ovChart').innerHTML = hChart(CATS.map((c, i) => ({
      label: c, value: n(rows[i].appDelta)})), {fmt: v => Bs(v, 2), L: 200});
    let h = '<thead><tr><th>Category</th><th class="num">Systems</th>'
      + '<th class="num">Base request</th><th class="num">Reconciliation request</th>'
      + '<th class="num">Total request</th><th class="num">Enacted NDAA</th>'
      + '<th class="num">Enacted appropriation</th><th class="num">Spend plan</th>'
      + '<th class="num">Enacted total</th><th class="num">Mandatory dependence</th></tr></thead><tbody>';
    rows.forEach(r => {
      h += '<tr><td class="name">' + esc(r.cat) + '</td><td class="num">' + N(r.n) + '</td>'
        + '<td class="num">' + B(n(r.base), 2) + '</td><td class="num">' + B(n(r.recon), 2) + '</td>'
        + '<td class="num">' + B(n(r.total), 2) + '</td><td class="num">' + B(n(r.ndaa), 2) + '</td>'
        + '<td class="num ' + cl(n(r.appDelta)) + '">' + B(n(r.app), 2) + '</td>'
        + '<td class="num">' + B(n(r.spend), 2) + '</td><td class="num">' + B(n(r.enactTotal), 2) + '</td>'
        + '<td class="num">' + P(n(r.mandDep)) + '</td></tr>';
    });
    h += '</tbody><tfoot><tr><td>All selected systems</td><td class="num">' + N(tot.n) + '</td>'
      + '<td class="num">' + B(n(tot.base), 2) + '</td><td class="num">' + B(n(tot.recon), 2) + '</td>'
      + '<td class="num">' + B(n(tot.total), 2) + '</td><td class="num">' + B(n(tot.ndaa), 2) + '</td>'
      + '<td class="num">' + B(n(tot.app), 2) + '</td><td class="num">' + B(n(tot.spend), 2) + '</td>'
      + '<td class="num">' + B(n(tot.enactTotal), 2) + '</td><td class="num">' + P(n(tot.mandDep)) + '</td>'
      + '</tr></tfoot>';
    $('#ovTable').innerHTML = h;
    $('#ovNote').textContent = 'Mandatory dependence is the share of enacted FY2026 resources that '
      + 'came from reconciliation rather than from the appropriation. Across all 97 systems it is '
      + P(n(tot.mandDep)) + '.';
  }

  /* largest lines and committee action */
  const yy = isBoth ? YEARS[2027] : y;
  const rws = isBoth ? YEARS[2027].rows : ROWS();
  const bf = yy.baselines[yy.defBase].f;
  const top = rws.slice().filter(d => has(yy.total(d)))
    .sort((a, b) => n(yy.total(b)) - n(yy.total(a))).slice(0, 14);
  $('#topLead').textContent = 'The fourteen largest lines by total request in ' + yy.label
    + ' carry ' + P(sum(top, d => n(yy.total(d))) / sum(rws, d => n(yy.total(d))))
    + ' of everything CRS tabulates.';
  $('#topChartT').textContent = 'Largest lines by total request, ' + yy.label;
  $('#topChart').innerHTML = hChart(top.map(d => ({label: d.name, value: n(yy.total(d))})),
    {fmt: v => B(v, 1), L: 230, w: 700});
  const akey = YEAR === 2027 || isBoth ? 'hacAction' : 'appAction';
  const alab = YEAR === 2027 || isBoth ? 'House appropriators' : 'Appropriators';
  const counts = {};
  rws.forEach(d => {const k = d[akey] || 'Not reported'; counts[k] = (counts[k] || 0) + 1;});
  const keys = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);
  $('#actChartT').textContent = 'What the ' + alab.toLowerCase() + ' did, by number of lines';
  $('#actChart').innerHTML = donut(keys.map((k, i) => ({label: k, value: counts[k], color: PAL[i % PAL.length]})),
    {w: 380, h: 300, center: String(rws.length), sub: 'lines', fmt: v => N(v) + ' lines'})
    + '<div style="margin-top:10px">' + keys.map((k, i) =>
      '<span class="tag" style="margin-right:8px"><span style="color:' + PAL[i % PAL.length]
      + '">&#9632;</span> ' + esc(k) + ' ' + counts[k] + '</span>').join('') + '</div>';

  /* context table */
  const ctxRows = (YEAR === 2027 && !isBoth)
    ? DATA.cats27.invest.map(r => ({item: r.item, amt: r.amt, basis: r.basis}))
    : DATA.ctx.topline.fy26.concat(DATA.ctx.topline.ctx)
        .map(r => ({item: r.item, amt: r.amt, basis: r.basis + (r.src ? '. ' + r.src : '')}));
  let ch = '<thead><tr><th>Item</th><th class="num">Amount</th><th>Basis</th></tr></thead><tbody>';
  ctxRows.forEach(r => {
    ch += '<tr><td class="name">' + esc(r.item) + '</td><td class="num">'
      + (has(r.amt) ? (Math.abs(n(r.amt)) < 1.2 && n(r.amt) !== 0 && n(r.amt) < 1 ? P(n(r.amt)) : B(n(r.amt), n(r.amt) >= 100 ? 1 : 2)) : BLANK)
      + '</td><td>' + esc(r.basis) + '</td></tr>';
  });
  $('#ctxTable').innerHTML = ch + '</tbody>';
}
"""

TEMPLATE_JS4 = r"""
/* ---------- scenario ---------- */
function renderScenario() {
  const R = compute(), rows = R.rows;
  const base = sum(rows, r => r.b), fin = sum(rows, r => r.final);
  const up = rows.filter(r => r.delta > 0.0005).length;
  const dn = rows.filter(r => r.delta < -0.0005).length;
  const net = fin - base;
  /* An offset rule can only draw on lines the instructions left alone. Where a
     whole category has been touched there is nothing left to draw on, so part of
     the net stays on the topline. That is said here rather than hidden. */
  let offNote;
  if (S.offset === 'none') offNote = 'the topline moves';
  else if (Math.abs(net) < 0.005) offNote = 'fully offset, the topline holds';
  else offNote = 'offset could not absorb ' + money(net) + ', no untouched line was left to draw on';
  $('#scenFigs').innerHTML = [
    fig(B(base, 2), 'Baseline', R.baseLabel),
    fig(B(fin, 2), 'After the instructions', Y().label),
    fig('<span class="' + cl(net) + '">' + Bs(net, 2) + '</span>', 'Net change', offNote),
    fig(N(up) + ' up, ' + N(dn) + ' down', 'Lines moved', N(rows.length) + ' lines in scope')
  ].join('');

  const ok = LEDGER.filter(e => e.ok);
  $('#ledger').innerHTML = LEDGER.length
    ? LEDGER.map(e => '<div class="li' + (e.ok ? '' : ' bad') + '">'
        + '<div class="rd">' + esc(e.ok ? e.reading : e.raw) + (e.ok ? ' <span class="mono ' + cl(e.dollars) + '">' + moneyS(e.dollars || 0) + '</span>' : '') + '</div>'
        + '<div class="mt">' + esc(e.ok
            ? ((e.purpose ? 'Purpose: ' + e.purpose + '. ' : '')
               + (e.offsetLabel ? 'Offset from ' + e.offsetLabel + '. ' : '')
               + (e.offsetMsg ? e.offsetMsg + ' ' : '')
               + (e.assumedUnit ? 'A bare number was read as billions. ' : '')
               + (e.scopeKind === 'category' || e.scopeKind === 'service' || e.scopeKind === 'all'
                  ? 'Spread across ' + e.ids.length + ' lines in proportion to their baseline.' : ''))
            : e.msg) + '</div></div>').join('')
    : '<p class="note" style="margin-top:0">Nothing written yet. One instruction a line, then press Price it.</p>';

  const catsD = CATS.map(c => sum(rows.filter(r => r.d.cat === c), r => r.delta));
  $('#scenChartT').textContent = 'Net change by category against ' + R.baseLabel.toLowerCase()
    + ', billions of dollars';
  $('#scenChart').innerHTML = hChart(CATS.map((c, i) => ({label: c, value: catsD[i]})),
    {fmt: v => Bs(v, 2), L: 210});

  TBL = 'sview';
  const moved = rows.filter(r => Math.abs(r.delta) > 0.0005);
  const show = (moved.length ? moved : rows.slice().sort((a, b) => b.b - a.b).slice(0, 20));
  const get = (r, k) => k === 'name' ? r.d.name : k === 'cat' ? r.d.cat : r[k];
  const sorted = sortRows(show, S.sview, get);
  let h = '<thead><tr>' + th('Line', 'name', 'name', true) + th('Category', 'cat', '', true)
    + th('Baseline', 'b', 'num', true) + th('After', 'final', 'num', true)
    + th('Change', 'delta', 'num', true) + th('Change (%)', 'pct', 'num', true)
    + '<th>Source of the change</th></tr></thead><tbody>';
  sorted.forEach(r => {
    const pc = r.b ? r.delta / r.b : null;
    const why = [];
    if (r.lp) why.push('written ' + Ps(r.lp));
    if (r.ld) why.push('written ' + moneyS(r.ld));
    if (S.global) why.push('across the board ' + Ps(S.global));
    if (r.off) why.push('offset ' + moneyS(r.off));
    h += '<tr><td class="name">' + esc(r.d.name) + '</td><td>' + esc(r.d.cat) + '</td>'
      + '<td class="num">' + B(r.b, 2) + '</td><td class="num">' + B(r.final, 2) + '</td>'
      + '<td class="num ' + cl(r.delta) + '">' + Bs(r.delta, 2) + '</td>'
      + '<td class="num ' + cl(r.delta) + '">' + (pc === null || !isFinite(pc) ? '<span class="dim">n/a</span>' : Ps(pc)) + '</td>'
      + '<td class="dim">' + esc(why.join(', ') || 'unchanged') + '</td></tr>';
  });
  h += '</tbody><tfoot><tr><td>' + (moved.length ? 'Lines that moved' : 'Twenty largest lines')
    + '</td><td></td><td class="num">' + B(sum(show, r => r.b), 2) + '</td>'
    + '<td class="num">' + B(sum(show, r => r.final), 2) + '</td>'
    + '<td class="num ' + cl(sum(show, r => r.delta)) + '">' + Bs(sum(show, r => r.delta), 2) + '</td>'
    + '<td></td><td></td></tr></tfoot>';
  $('#scenTable').innerHTML = h;
  $('#scenTable').dataset.view = 'sview';
}

function scenarioMemo() {
  const R = compute(), rows = R.rows;
  const base = sum(rows, r => r.b), fin = sum(rows, r => r.final);
  const out = [Y().label + ' scenario memo',
               'Baseline: ' + R.baseLabel + ', ' + B(base, 2),
               'Source: ' + Y().report, ''];
  const ok = (LEDGER || []).filter(e => e.ok);
  if (ok.length) {
    out.push('Written instructions');
    ok.forEach((e, i) => out.push('  ' + (i + 1) + '. ' + e.reading + '. '
      + (e.purpose ? 'Purpose: ' + e.purpose + '. ' : '')
      + (e.offsetLabel ? 'Offset from ' + e.offsetLabel + '. ' : '')
      + 'Effect: ' + moneyS(e.dollars || 0) + '.'));
    out.push('');
  }
  if (S.global) out.push('Across the board: ' + Ps(S.global) + '.');
  if (S.offset !== 'none') out.push('Offset rule: ' + ({
    all: 'pro rata across untouched lines',
    cat: 'each category held constant',
    sys: 'absorbed by ' + S.offsetSys}[S.offset]) + '.');
  out.push('Net change: ' + Bs(fin - base, 2) + '. Scenario total: ' + B(fin, 2) + '.');
  const one = payCompute().one;
  if (one > 0) out.push('For scale, that net change equals a ' + Ps((fin - base) / one * 0.01, 2)
    + ' across-the-board military basic pay raise for one year, fully burdened, priced on the '
    + 'FY' + PAYYEAR() + ' rates.');
  out.push('');
  out.push('Every figure is read from ' + DATA.meta.workbook + '. Nothing was recomputed.');
  return out.join('\n');
}

/* ---------- claim tester ---------- */
function renderClaim() {
  const c = claimCompute();
  if (!c) {$('#claimBody').innerHTML = '<p class="note">Choose a line.</p>'; return;}
  const y = Y();
  const figs = [
    fig('<span class="' + cl(c.chg) + '">' + amountWords(c.chg) + '</span>', 'The change',
        c.own ? 'a claim of your own' : c.baseLabel),
    c.own
      ? fig(esc(c.label), 'What the money is for', 'not a line in either report', true)
      : fig(B(c.b, 2) + ' to ' + B(c.nw, 2), 'That line, before and after', c.label),
    c.own
      ? fig(c.allTotal ? Ps(c.chg / c.allTotal, 4) : 'n/a', 'As a share of every selected line',
            B(c.allTotal, 1) + ' on this baseline')
      : fig(c.b ? Ps(c.chg / c.b) : 'n/a', 'As a share of its own line', esc(c.d.cat)),
    fig(c.raise === null ? 'n/a' : Ps(c.raise, 4), 'As an across-the-board pay raise',
        'FY' + PAYYEAR() + ' rates, fully burdened')
  ];
  let h = '<div class="figrow" style="margin-top:0">' + figs.join('') + '</div>';
  h += '<div class="panel"><h4 style="margin-top:0">What it buys, and what it costs elsewhere</h4>';
  h += '<div class="tablewrap"><table><tbody>';
  const row = (k, v, note) => '<tr><td class="name">' + esc(k) + '</td><td class="num">' + v
    + '</td><td class="dim">' + esc(note || '') + '</td></tr>';

  h += row('The amount, in dollars', '$' + N(Math.abs(c.chg) * 1e9),
    'the same figure as ' + B(c.chg, c.chg !== 0 && Math.abs(c.chg) < 0.01 ? 6 : 2));

  if (!c.own) {
    h += row('Budget-line cost per unit', c.uc === null ? BLANK : '$' + N(c.uc, 1) + 'M',
      c.uc === null ? 'CRS printed no quantity for this line, so no unit cost can be formed'
        : 'the whole budget line divided by the quantity CRS printed, not a flyaway cost');
    h += row('Units at that rate', c.units === null ? BLANK : N(c.units, 2),
      c.units === null ? '' : 'indicative only, because a budget line is not linear in quantity');
  }
  if (c.ref) {
    h += row('The same money in ' + c.ref.name,
      c.refUnits === null ? BLANK : N(c.refUnits, c.refUnits < 1 ? 4 : 2) + ' units',
      c.refUC === null ? 'no published quantity for the comparison line'
        : 'at $' + N(c.refUC, 1) + 'M a unit');
  }
  if (!c.own) {
    h += row('Its share of ' + c.d.cat, c.catTotal ? P(c.b / c.catTotal) : 'n/a',
      'category total ' + B(c.catTotal, 2) + ' on this baseline');
  }
  h += row(c.own ? 'The claim against every selected line' : 'Its share of every selected line',
    c.allTotal ? Ps((c.own ? c.chg : c.b) / c.allTotal, 4) : 'n/a',
    'portfolio total ' + B(c.allTotal, 2));
  if (c.pool) {
    h += row('Paid for by ' + c.poolLabel, c.pctCut === null ? 'n/a' : Ps(c.pctCut, 4),
      'a pool of ' + B(c.pool, 2) + ', so the offset is ' + Ps(c.pctCut, 4) + ' across it');
  } else if (S.claim.off !== 'none') {
    h += row('Paid for by ' + (c.poolLabel || 'that source'), BLANK,
      'that source carries no money on this baseline');
  } else {
    h += row('Offset', '<span class="dim">none</span>',
      'the topline moves by the full amount');
  }
  h += '</tbody></table></div>';

  const pay = payCompute();
  const perMember = (Math.abs(c.chg) * 1e9) / (sum(GRADES.map(g => strengthFor(g, 'All services')), x => x) || 1);
  h += '<p class="note" style="margin-top:14px">Read in one sentence: '
    + (c.own
      ? esc(c.label) + ' at ' + amountWords(c.chg) + ' is ' + Ps(c.chg / c.allTotal, 5)
        + ' of the ' + B(c.allTotal, 1) + ' CRS tabulates for ' + y.label
        + (c.pool ? ', would come to ' + Ps(c.pctCut, 4) + ' of ' + esc(c.poolLabel) : ', unoffset')
        + ', and is worth about ' + (c.raise === null ? 'an unpriced fraction of'
            : Ps(c.raise, 5) + ' of') + ' a one year across-the-board military pay raise. '
        + 'Spread across the active-duty force it is ' + USD(perMember) + ' a member.'
      : 'a change of ' + amountWords(c.chg) + ' to ' + esc(c.label) + ' is '
        + (c.b ? Ps(c.chg / c.b) : 'an unpriced change') + ' of that line, '
        + (c.pool ? Ps(c.pctCut, 2) + ' of ' + esc(c.poolLabel) : 'unoffset')
        + ', and about ' + (c.raise === null ? 'an unpriced fraction of' : Ps(c.raise, 2) + ' of')
        + ' a one year across-the-board military pay raise.')
    + '</p></div>';
  if (!c.own) h += '<p class="note">' + esc(c.d.desc || c.d.footnote || '') + '</p>';
  else h += '<p class="note">A claim of your own is measured against the same figures as every '
    + 'line on this page, but it is not one of them. It has no baseline, no published quantity '
    + 'and no unit cost, so those rows stay blank rather than being invented.</p>';
  $('#claimBody').innerHTML = h;
}

/* ---------- pay ---------- */
function renderPay() {
  const P0 = payCompute(), R = P0.rates;
  $('#payLead').textContent = 'Basic pay comes from the January 2026 table. The burden rates come '
    + 'from the M-1 for the year being priced: retired pay accrual at ' + P(R.rpa, 1)
    + ' of basic pay, agency contributions to the Thrift Savings Plan at '
    + P(R.tspAuto + R.tspMatch, 1) + ', and employer payroll taxes on top. '
    + 'A raise moves basic pay. It does not move allowances, so it does not move the whole of '
    + 'regular military compensation.';
  $('#payFigs').innerHTML = [
    fig(B(P0.total, 2), 'Full-year burdened cost', 'FY' + PAYYEAR() + ' rates, ' + esc(P0.svc)),
    fig(B(P0.first, 2), 'Cost in the first fiscal year', N(P0.months) + ' of 12 months'),
    fig(N(P0.affected), 'Members receiving a raise', 'of ' + N(sum(GRADES.map(g => strengthFor(g, P0.svc)), x => x))),
    fig(P0.affected ? USD(P0.avg) : 'n/a', 'Average basic pay increase a member', 'before burdens')
  ].join('');
  $('#payChartT').textContent = 'Burdened cost of the raise by pay grade, millions of dollars';
  $('#payChart').innerHTML = vChart(GRADES, [
    {name: 'Basic pay', color: '#a8506e', data: P0.rows.map(r => r.basic * 1000)},
    {name: 'Retired pay accrual', color: '#d98ba3', data: P0.rows.map(r => r.rpa * 1000)},
    {name: 'Payroll taxes', color: '#c9b6d9', data: P0.rows.map(r => r.fica * 1000)},
    {name: 'Thrift Savings Plan', color: '#4a7f92', data: P0.rows.map(r => r.tsp * 1000)}
  ], {stack: true, h: 320, fmt: v => N(v, 0)});

  let h = '<thead><tr><th>Grade</th><th class="num">Strength</th><th class="num">Years of service</th>'
    + '<th class="num">Monthly basic pay</th><th class="num">Raise</th>'
    + '<th class="num">New monthly pay</th><th class="num">Annual increase a member</th>'
    + '<th class="num">Basic pay cost</th><th class="num">Burdened cost</th>'
    + '<th class="num">First-year cost</th></tr></thead><tbody>';
  P0.rows.forEach(r => {
    h += '<tr><td class="name">' + esc(r.g) + '</td><td class="num">' + N(r.strength) + '</td>'
      + '<td class="num">' + esc(REPYOS[r.g]) + '</td>'
      + '<td class="num">' + USD(r.pay) + '</td>'
      + '<td class="num ' + (r.tot ? 'pos' : 'dim') + '">' + P(r.tot, 1) + '</td>'
      + '<td class="num">' + USD(r.nw) + (r.nw >= r.cap - 0.01 && r.tot ? ' <span class="tag chk">at cap</span>' : '') + '</td>'
      + '<td class="num">' + USD(r.annual) + '</td>'
      + '<td class="num">$' + N(r.basic * 1000, 1) + 'M</td>'
      + '<td class="num">$' + N(r.total * 1000, 1) + 'M</td>'
      + '<td class="num">$' + N(r.first * 1000, 1) + 'M</td></tr>';
  });
  h += '</tbody><tfoot><tr><td>All grades</td><td class="num">' + N(sum(P0.rows, r => r.strength))
    + '</td><td></td><td></td><td></td><td></td><td></td>'
    + '<td class="num">' + B(P0.basic, 2) + '</td><td class="num">' + B(P0.total, 2) + '</td>'
    + '<td class="num">' + B(P0.first, 2) + '</td></tr></tfoot>';
  $('#payTable').innerHTML = h;
  $('#payNote').textContent = 'A grade whose pay is already at the statutory monthly cap cannot '
    + 'take the full raise, so its cost is the distance to the cap and no more. One percentage '
    + 'point across the whole active-duty force, fully burdened, costs ' + B(P0.one, 2)
    + ' on FY' + PAYYEAR() + ' rates.';

  /* RMC */
  const AL = PAY.allow, W = AL.weighted;
  $('#rmcLead').textContent = 'Regular military compensation is basic pay, the average housing '
    + 'allowance, the subsistence allowance, and the federal income tax advantage that follows '
    + 'from the two allowances not being taxed. The published Selected Military Compensation '
    + 'Tables are dated 1 January 2023, so the 2023 structure supplies the shape and the M-1 for '
    + 'the priced year supplies the level. The housing figure is the amount per member drawing '
    + 'cash housing, not a published allowance rate, because the M-1 line includes the overseas '
    + 'housing allowance.';
  let rh = '<thead><tr><th>Grade</th><th class="num">Strength</th>'
    + '<th class="num">Basic pay</th><th class="num">Housing</th><th class="num">Subsistence</th>'
    + '<th class="num">Tax advantage</th><th class="num">Regular military compensation</th>'
    + '<th class="num">Basic pay share</th><th class="num">Published RMC, 2023</th>'
    + '</tr></thead><tbody>';
  AL.rows.forEach(r => {
    rh += '<tr><td class="name">' + esc(r.g) + '</td><td class="num">' + N(r.strength) + '</td>'
      + '<td class="num">' + USD(r.basicYr) + '</td><td class="num">' + USD(r.bahYr) + '</td>'
      + '<td class="num">' + USD(r.basYr) + '</td><td class="num">' + USD(r.txad) + '</td>'
      + '<td class="num">' + USD(r.rmc) + '</td><td class="num">' + P(n(r.basicShare)) + '</td>'
      + '<td class="num dim">' + USD(r.pubRmc) + '</td></tr>';
  });
  rh += '</tbody>';
  $('#rmcTable').innerHTML = rh;
  $('#rmcNote').textContent = 'Strength-weighted, basic pay is ' + P(n(W.sharePriced))
    + ' of regular military compensation on the priced year against ' + P(n(W.sharePub))
    + ' in the published 2023 tables, and the implied marginal tax rate behind the tax advantage '
    + 'is ' + P(n(W.ratePriced), 1) + ' against ' + P(n(W.ratePub), 1) + '. '
    + AL.recon.map(r => r.check + ': ' + (Math.abs(n(r.value) - n(r.target)) < 0.5
        ? 'reconciles exactly' : 'differs by ' + N(n(r.value) - n(r.target), 3))).join('. ') + '.';

  /* composite comparison */
  const CP = PAY.composite, svcC = A['Service for the composite rate comparison'] || 'Army';
  const blk = PAYYEAR() === 2026 ? CP.fy26 : CP.fy27;
  let ch = '<thead><tr><th>Grade</th><th class="num">Model burdened cost a member</th>'
    + '<th class="num">Composite standard rate, ' + esc(svcC) + '</th>'
    + '<th class="num">Spread</th><th class="num">Spread (%)</th>'
    + '<th>What the difference is</th></tr></thead><tbody>';
  PAY.grades.forEach(g => {
    const cmp = (blk[g.g] || {})[svcC];
    const cv = cmp ? cmp.comp : null;
    const model = n(g.model);
    const sp = has(cv) ? model - n(cv) : null;
    ch += '<tr><td class="name">' + esc(g.g) + '</td>'
      + '<td class="num">' + USD(model) + '</td>'
      + '<td class="num">' + (has(cv) ? USD(n(cv)) : BLANK) + '</td>'
      + '<td class="num ' + (sp === null ? 'dim' : cl(sp)) + '">' + (sp === null ? BLANK : USD(sp)) + '</td>'
      + '<td class="num ' + (sp === null ? 'dim' : cl(sp)) + '">' + (sp === null || !cv ? BLANK : Ps(sp / n(cv))) + '</td>'
      + '<td class="dim">' + (sp === null ? 'the service published no rate for this grade'
          : 'the composite rate carries allowances, special pay and permanent change of station costs that the model leaves out') + '</td></tr>';
  });
  ch += '</tbody><tfoot><tr><td>Strength weighted</td><td class="num">' + USD(n(PAY.weighted.model))
    + '</td><td class="num">' + USD(n(PAY.weighted.composite)) + '</td>'
    + '<td class="num">' + USD(n(PAY.weighted.model) - n(PAY.weighted.composite)) + '</td>'
    + '<td class="num">' + Ps((n(PAY.weighted.model) - n(PAY.weighted.composite)) / n(PAY.weighted.composite)) + '</td>'
    + '<td class="dim">' + esc(A['Composite rate edition in use'] || '') + '</td></tr></tfoot>';
  $('#compTable').innerHTML = ch;
}
"""

TEMPLATE_JS5 = r"""
/* ---------- projection ---------- */
function renderProjection() {
  const rows = projCompute();
  const cum = sum(rows, r => r.cum), first = sum(rows, r => r.gaps[0]);
  const bal = n(DATA.ctx.proj.balance);
  $('#projLead').textContent = 'DOD allocated almost all of its reconciliation money in FY2026. '
    + 'Those dollars do not repeat. This holds the FY2026 programme constant and asks what '
    + 'discretionary appropriations would have to carry from FY2027 onward to sustain it, before '
    + 'any new content. It is arithmetic on a stated assumption, not a forecast.';
  $('#projFigs').innerHTML = [
    fig(B(first, 1), 'Gap in FY2027', 'first year without reconciliation money'),
    fig(B(cum, 1), 'Cumulative FY2027 to FY2031', 'five years at the stated growth'),
    fig(P(n(DATA.cats26.total.mandDep)), 'Mandatory dependence, FY2026 enacted',
        B(n(DATA.cats26.total.spend), 1) + ' of ' + B(n(DATA.cats26.total.enactTotal), 1)),
    fig(B(bal, 1), 'Unallocated P.L. 119-21 balance', 'available to offset the first year')
  ].join('');
  $('#projChartT').textContent = 'Sustainment gap by category, billions of dollars a year';
  $('#projChart').innerHTML = vChart(CATS, [1, 2, 3, 4, 5].map((t, i) => ({
    name: 'FY' + (2026 + t), color: PAL[i % PAL.length],
    data: rows.map(r => r.gaps[i])
  })), {h: 320, fmt: v => v.toFixed(0)});
  let h = '<thead><tr><th>Category</th><th class="num">FY2026 appropriation</th>'
    + '<th class="num">FY2026 spend plan</th><th class="num">Mandatory dependence</th>'
    + '<th class="num">Discretionary growth needed</th>'
    + '<th class="num">FY2027</th><th class="num">FY2028</th><th class="num">FY2029</th>'
    + '<th class="num">FY2030</th><th class="num">FY2031</th><th class="num">Cumulative</th>'
    + '</tr></thead><tbody>';
  rows.forEach(r => {
    h += '<tr><td class="name">' + esc(r.c) + '</td><td class="num">' + B(r.app, 2) + '</td>'
      + '<td class="num">' + B(r.spend, 2) + '</td><td class="num">' + P(r.dep) + '</td>'
      + '<td class="num">' + (r.growth === null ? BLANK : Ps(r.growth)) + '</td>'
      + r.gaps.map(g => '<td class="num">' + B(g, 2) + '</td>').join('')
      + '<td class="num">' + B(r.cum, 2) + '</td></tr>';
  });
  h += '</tbody><tfoot><tr><td>All selected systems</td>'
    + '<td class="num">' + B(sum(rows, r => r.app), 2) + '</td>'
    + '<td class="num">' + B(sum(rows, r => r.spend), 2) + '</td>'
    + '<td class="num">' + P(sum(rows, r => r.spend) / (sum(rows, r => r.tot) || 1)) + '</td>'
    + '<td class="num">' + Ps(sum(rows, r => r.tot) * (1 + n(S.proj.infl)) * (1 + n(S.proj.real)) / (sum(rows, r => r.app) || 1) - 1) + '</td>'
    + [0, 1, 2, 3, 4].map(i => '<td class="num">' + B(sum(rows, r => r.gaps[i]), 2) + '</td>').join('')
    + '<td class="num">' + B(cum, 2) + '</td></tr></tfoot>';
  $('#projTable').innerHTML = h;
  $('#projNote').textContent = 'The unallocated balance of ' + B(bal, 1) + ' from P.L. 119-21 '
    + 'stays available into FY2027, so the first-year gap net of it is about '
    + B(Math.max(first - bal, 0), 1) + '. Replacement is set to ' + P(n(S.proj.repl), 0)
    + ', so ' + (n(S.proj.repl) >= 0.999 ? 'the model assumes Congress backfills every reconciliation dollar and the gap closes.'
      : n(S.proj.repl) <= 0.001 ? 'no reconciliation dollar is assumed to be replaced.'
      : 'that share of the reconciliation money is assumed replaced.');
}

/* ---------- year on year ---------- */
function renderCompare() {
  const C = DATA.compare;
  $('#cmpNotes').innerHTML = C.notes.map((t, i) => '<p class="' + (i === 0 ? 'body' : 'note')
    + '" style="margin:' + (i === 0 ? '0 0 10px' : '4px 0 0') + '">' + esc(t) + '</p>').join('');
  const both = C.rows.filter(r => r.presence === 'Both years');
  /* Listed in both reports is not the same as carrying a figure in both. Four
     lines are listed in R49023 with no FY2027 figure printed against them, and
     a line with nothing printed cannot be differenced. The bridge below uses the
     stricter test, so these figures use it too. */
  const priced = both.filter(r => has(r.tot26) && has(r.tot27));
  const matched26 = sum(priced, r => n(r.tot26)), matched27 = sum(priced, r => n(r.tot27));
  const listedNotPriced = both.filter(r => !has(r.tot26) || !has(r.tot27));
  const newOnly = C.only.filter(o => /FY2027 only/.test(o.dir));
  const gone = C.only.filter(o => /FY2026 only/.test(o.dir));
  $('#cmpFigs').innerHTML = [
    fig(N(both.length), 'Lines listed in both reports', 'the joined spine'),
    fig(N(priced.length), 'Lines carrying a figure in both',
        listedNotPriced.length + ' listed with no FY2027 figure printed'),
    fig(Bs(matched27 - matched26, 1), 'Change on the lines that can be differenced',
        B(matched26, 1) + ' to ' + B(matched27, 1)),
    fig(N(newOnly.length) + ' new, ' + N(gone.length) + ' dropped', 'Lines in one report only',
        B(sum(newOnly, o => n(o.tot27)), 2) + ' new, ' + B(sum(gone, o => n(o.tot26)), 2) + ' dropped')
  ].join('');
  if (listedNotPriced.length) {
    $('#cmpFigs').insertAdjacentHTML('afterend',
      '<p class="note" id="cmpBlankNote">' + esc(listedNotPriced.map(r => r.name).join(', '))
      + ' appear in both reports, but R49023 prints no FY2027 figure against them. '
      + 'They stay on the spine and stay blank. A programme CRS did not price is not a programme cut to zero.'
      + '</p>');
    const old = document.getElementById('cmpBlankNote');
    if (old && old.previousElementSibling && old.previousElementSibling.id !== 'cmpFigs') old.remove();
  }
  const movers = both.slice().filter(r => has(r.dTot))
    .sort((a, b) => Math.abs(n(b.dTot)) - Math.abs(n(a.dTot))).slice(0, 16);
  $('#cmpChartT').textContent = 'Largest movements on the joined spine, total request, billions of dollars';
  $('#cmpChart').innerHTML = hChart(movers.map(r => ({label: r.name, value: n(r.dTot)})),
    {fmt: v => Bs(v, 2), L: 250, w: 760});

  let bh = '<thead><tr><th>Step</th><th class="num">Amount</th><th>Note</th></tr></thead><tbody>';
  C.bridge.forEach(b => {
    bh += '<tr><td class="name">' + esc(b.step) + '</td><td class="num">'
      + (has(b.amt) ? B(n(b.amt), 2) : BLANK) + '</td><td class="dim">' + esc(b.note) + '</td></tr>';
  });
  $('#bridgeTable').innerHTML = bh + '</tbody>';

  let oh = '<thead><tr><th>Direction</th><th>Category</th><th>Line</th>'
    + '<th class="num">FY2026 total request</th><th class="num">FY2027 total request</th>'
    + '</tr></thead><tbody>';
  C.only.forEach(o => {
    oh += '<tr><td>' + esc(o.dir) + '</td><td>' + esc(o.cat) + '</td>'
      + '<td class="name">' + esc(o.sys) + '</td>'
      + '<td class="num">' + cellB(o.tot26, v => B(n(v), 2)) + '</td>'
      + '<td class="num">' + cellB(o.tot27, v => B(n(v), 2)) + '</td></tr>';
  });
  $('#onlyTable').innerHTML = oh + '</tbody>';

  TBL = 'cview';
  const v = S.cview;
  const q = (v.q || '').toLowerCase();
  let rows = C.rows.filter(r => {
    if (v.cat !== 'all' && r.cat !== v.cat) return false;
    if (v.pres !== 'all' && r.presence !== v.pres) return false;
    if (q && (String(r.name) + ' ' + r.id26 + ' ' + r.id27 + ' ' + r.alt).toLowerCase().indexOf(q) < 0) return false;
    return true;
  });
  rows = sortRows(rows, v, (r, k) => r[k]);
  let h = '<thead><tr>' + th('Line', 'name', 'name', true) + th('Category', 'cat', '', true)
    + th('Present in', 'presence', '', true)
    + th('FY2026 base', 'base26', 'num', true) + th('FY2027 discretionary', 'disc27', 'num', true)
    + th('Change', 'dDisc', 'num', true)
    + th('FY2026 total', 'tot26', 'num', true) + th('FY2027 total', 'tot27', 'num', true)
    + th('Change', 'dTot', 'num', true) + th('Change (%)', 'dTotPct', 'num', true)
    + '<th>Crosswalk basis</th></tr></thead><tbody>';
  rows.forEach(r => {
    h += '<tr><td class="name">' + esc(r.name)
      + (r.alt ? '<div class="dim" style="font-size:12px">FY2026: ' + esc(r.alt) + '</div>' : '')
      + '</td><td>' + esc(r.cat) + '</td><td>' + esc(r.presence) + '</td>'
      + '<td class="num">' + cellB(r.base26, x => B(n(x), 2)) + '</td>'
      + '<td class="num">' + cellB(r.disc27, x => B(n(x), 2)) + '</td>'
      + '<td class="num ' + cl(n(r.dDisc)) + '">' + cellB(r.dDisc, x => Bs(n(x), 2)) + '</td>'
      + '<td class="num">' + cellB(r.tot26, x => B(n(x), 2)) + '</td>'
      + '<td class="num">' + cellB(r.tot27, x => B(n(x), 2)) + '</td>'
      + '<td class="num ' + cl(n(r.dTot)) + '">' + cellB(r.dTot, x => Bs(n(x), 2)) + '</td>'
      + '<td class="num ' + cl(n(r.dTotPct)) + '">' + cellB(r.dTotPct, x => Ps(n(x))) + '</td>'
      + '<td class="dim">' + esc(r.basis) + '</td></tr>';
  });
  h += '</tbody><tfoot><tr><td>' + N(rows.length) + ' lines shown</td><td></td><td></td>'
    + '<td class="num">' + B(sum(rows, r => n(r.base26)), 2) + '</td>'
    + '<td class="num">' + B(sum(rows, r => n(r.disc27)), 2) + '</td>'
    + '<td class="num">' + Bs(sum(rows, r => n(r.dDisc)), 2) + '</td>'
    + '<td class="num">' + B(sum(rows, r => n(r.tot26)), 2) + '</td>'
    + '<td class="num">' + B(sum(rows, r => n(r.tot27)), 2) + '</td>'
    + '<td class="num">' + Bs(sum(rows, r => n(r.dTot)), 2) + '</td>'
    + '<td></td><td></td></tr></tfoot>';
  $('#cmpTable').innerHTML = h;
  $('#cmpTable').dataset.view = 'cview';
}

/* ---------- the data ---------- */
function dataCols() {
  const y = Y();
  const base = [
    {k: 'id', label: 'ID', cls: 'id'},
    {k: 'name', label: 'Line', cls: 'name'},
    {k: 'cat', label: 'Category', cls: ''},
    {k: 'svc', label: 'Service', cls: ''}
  ];
  if (YEAR === 2027) {
    return base.concat([
      {k: 'disc', label: 'Discretionary request', cls: 'num', m: true},
      {k: 'mand', label: 'Mandatory request', cls: 'num', m: true},
      {k: 'total', label: 'Total request', cls: 'num', m: true},
      {k: 'house', label: 'H.R. 8800 House', cls: 'num', m: true},
      {k: 'sasc', label: 'S. 4784 SASC', cls: 'num', m: true},
      {k: 'hac', label: 'H.R. 9495 HAC', cls: 'num', m: true},
      {k: 'hacDelta', label: 'HAC against request', cls: 'num', m: true, s: true},
      {k: 'unit', label: 'Cost per unit ($M)', cls: 'num', u: true},
      {k: 'enactApp', label: 'Enacted', cls: ''},
      {k: 'page', label: 'CRS page', cls: 'num', p: true}
    ]);
  }
  return base.concat([
    {k: 'base', label: 'Base request', cls: 'num', m: true},
    {k: 'recon', label: 'Reconciliation request', cls: 'num', m: true},
    {k: 'total', label: 'Total request', cls: 'num', m: true},
    {k: 'ndaa', label: 'Enacted NDAA', cls: 'num', m: true},
    {k: 'app', label: 'Enacted appropriation', cls: 'num', m: true},
    {k: 'spend', label: 'Spend plan', cls: 'num', m: true},
    {k: 'enactTotal', label: 'Enacted total', cls: 'num', m: true},
    {k: 'appDelta', label: 'Appropriation against base', cls: 'num', m: true, s: true},
    {k: 'unit', label: 'Cost per unit ($M)', cls: 'num', u: true},
    {k: 'appAction', label: 'Appropriators', cls: ''}
  ]);
}
function filtered() {
  const v = S.dview, q = (v.q || '').toLowerCase();
  return ROWS().filter(d => {
    if (v.cat !== 'all' && d.cat !== v.cat) return false;
    if (v.svc !== 'all' && d.svc !== v.svc) return false;
    if (q && (d.name + ' ' + d.id + ' ' + d.cat + ' ' + (d.svc || '') + ' '
              + (d.section || d.sub || '')).toLowerCase().indexOf(q) < 0) return false;
    return true;
  });
}
function renderData() {
  TBL = 'dview';
  const y = Y(), cols = dataCols();
  $('#dTitle').textContent = 'Every line CRS printed for ' + (S.year === 'both' ? 'FY2027' : y.label);
  $('#dLead').textContent = y.label + ' lines as published in ' + y.report
    + '. A blank cell means CRS printed no figure on that line.';
  let rows = filtered();
  if (!S.dview.sort || !cols.some(c => c.k === S.dview.sort)) S.dview.sort = cols[4].k;
  rows = sortRows(rows, S.dview, (d, k) => d[k]);
  let h = '<thead><tr>' + cols.map(c => th(c.label, c.k, c.cls, true)).join('') + '</tr></thead><tbody>';
  rows.forEach(d => {
    h += '<tr>' + cols.map(c => {
      const v = d[c.k];
      let out;
      if (c.m) out = has(v) ? (c.s ? '<span class="' + cl(n(v)) + '">' + Bs(n(v), 2) + '</span>' : B(n(v), 2)) : BLANK;
      else if (c.u) out = has(v) ? '$' + N(n(v), 1) : BLANK;
      else if (c.p) out = has(v) ? N(n(v)) : BLANK;
      else if (c.k === 'name') out = esc(v) + (d.printedAs ? ' <span class="tag chk">' + esc(d.printedAs) + '</span>' : '')
        + (d.inCat === 'No' ? ' <span class="tag">appendix</span>' : '');
      else if (c.k === 'enactApp') out = (v === 'Pending' || !has(v))
        ? '<span class="pending">pending</span>' : esc(v);
      else out = has(v) ? esc(v) : '<span class="dim">not classified</span>';
      return '<td class="' + (c.cls || '') + '">' + out + '</td>';
    }).join('') + '</tr>';
  });
  const money = cols.filter(c => c.m);
  h += '</tbody><tfoot><tr><td></td><td>' + N(rows.length) + ' lines</td><td></td><td></td>'
    + money.map(c => '<td class="num">' + (c.s ? Bs(sum(rows, d => n(d[c.k])), 2) : B(sum(rows, d => n(d[c.k])), 2)) + '</td>').join('')
    + '<td></td><td></td>' + (YEAR === 2027 ? '<td></td>' : '') + '</tr></tfoot>';
  $('#dTable').innerHTML = h;
  $('#dTable').dataset.view = 'dview';
  $('#dCount').textContent = N(rows.length) + ' of ' + N(ROWS().length) + ' lines';
}
function csvData() {
  const cols = dataCols(), rows = filtered();
  const head = cols.map(c => c.label).join(',');
  const body = rows.map(d => cols.map(c => {
    const v = d[c.k];
    if (v === null || v === undefined || v === '') return '';
    return /[",\n]/.test(String(v)) ? '"' + String(v).replace(/"/g, '""') + '"' : String(v);
  }).join(',')).join('\n');
  return head + '\n' + body + '\n';
}
function csvScenario() {
  const R = compute();
  const head = 'ID,Line,Category,Service,Baseline,Scenario,Change,Change (%)';
  const body = R.rows.map(r => [r.d.id, '"' + String(r.d.name).replace(/"/g, '""') + '"', r.d.cat,
    r.d.svc || '', r.b.toFixed(4), r.final.toFixed(4), r.delta.toFixed(4),
    r.b ? (r.delta / r.b).toFixed(6) : ''].join(',')).join('\n');
  return head + '\n' + body + '\n';
}
"""

TEMPLATE_JS6 = r"""
/* ---------- method ---------- */
function renderMethod() {
  const Q = DATA.ctx;
  let h = '<thead><tr><th>Category</th><th>Measure</th><th class="num">This workbook</th>'
    + '<th class="num">CRS narrative</th><th class="num">Difference</th><th>Status</th>'
    + '<th>Where</th><th>Explanation</th></tr></thead><tbody>';
  Q.qa26.forEach(r => {
    const st = r.status === 'PASS' ? 'ok' : 'rev';
    h += '<tr><td>' + esc(r.cat) + '</td><td class="name">' + esc(r.measure) + '</td>'
      + '<td class="num">' + (has(r.book) ? N(n(r.book), 2) : BLANK) + '</td>'
      + '<td class="num">' + (has(r.crs) ? N(n(r.crs), 2) : BLANK) + '</td>'
      + '<td class="num ' + cl(n(r.diff)) + '">' + (has(r.diff) ? n(r.diff).toFixed(2) : BLANK) + '</td>'
      + '<td><span class="tag ' + st + '">' + esc(r.status) + '</span></td>'
      + '<td class="dim">' + esc(r.where) + '</td><td class="dim">' + esc(r.why) + '</td></tr>';
  });
  h += '</tbody><tfoot><tr><td colspan="8">'
    + esc(Q.qa26.filter(r => r.status === 'PASS').length + ' pass, '
      + Q.qa26.filter(r => r.status === 'REVIEW').length
      + ' marked for review because the CRS narrative disagrees with the CRS tables, and '
      + Q.qa26struct.length + ' structural checks all pass.')
    + '</td></tr></tfoot>';
  $('#qaTable').innerHTML = h;

  let q = '<thead><tr><th>Section</th><th>Check</th><th class="num">This workbook</th>'
    + '<th class="num">Expected</th><th class="num">Difference</th><th>Status</th><th>Note</th>'
    + '</tr></thead><tbody>';
  Q.qa27.forEach(r => {
    q += '<tr><td class="dim">' + esc(r.section.slice(0, 2)) + '</td>'
      + '<td class="name">' + esc(r.check) + '</td>'
      + '<td class="num">' + (typeof r.book === 'number' ? N(r.book, Math.abs(r.book) < 10 ? 3 : 2) : esc(r.book)) + '</td>'
      + '<td class="num">' + (has(r.expect) ? N(n(r.expect), Math.abs(n(r.expect)) < 10 ? 3 : 2) : BLANK) + '</td>'
      + '<td class="num ' + cl(n(r.diff)) + '">' + (has(r.diff) ? n(r.diff).toFixed(4) : BLANK) + '</td>'
      + '<td><span class="tag ' + (r.status === 'OK' ? 'ok' : 'chk') + '">' + esc(r.status) + '</span></td>'
      + '<td class="dim">' + esc(r.note) + '</td></tr>';
  });
  q += '</tbody><tfoot><tr><td colspan="7">' + esc(Q.qa27counts.ok + ' checks pass and '
    + Q.qa27counts.check + ' is marked CHECK: the CRS summary says the Senate Armed Services '
    + 'Committee proposed authorizing the amount requested, while the committee columns in the '
    + 'same report net to a different figure. Both are shown. Neither is adjusted.')
    + '</td></tr></tfoot>';
  $('#qa27Table').innerHTML = q;

  const prose = Q.qa27sections.filter(s => s.notes.length);
  $('#contra').innerHTML = prose.map(s => '<h5 style="margin:18px 0 6px;font-weight:400">'
    + esc(s.title) + '</h5><div class="tablewrap"><table><tbody>'
    + s.notes.map(nn => '<tr><td class="name">' + esc(nn.check) + '</td><td>'
      + esc(typeof nn.book === 'number' ? String(nn.book) : nn.book) + '</td></tr>').join('')
    + '</tbody></table></div>').join('');

  /* Both sheets carry the same columns and one continuous reference numbering,
     so they are presented as a single list. */
  let sh = '<thead><tr><th class="id">Ref</th><th>Source</th><th>Author or publisher</th>'
    + '<th>Date</th><th>What it supplies</th></tr></thead><tbody>';
  Q.sources26.concat(Q.sources27).forEach(r => {
    const t = esc(r.title);
    sh += '<tr><td class="id">' + esc(r.id) + '</td>'
      + '<td class="name">'
      + (r.link ? '<a href="' + esc(r.link) + '" rel="noopener noreferrer" target="_blank">'
          + t + '</a>' : t)
      + '</td><td>' + esc(r.author) + '</td>'
      + '<td>' + esc(r.date) + '</td><td class="dim">' + esc(r.supplies) + '</td></tr>';
  });
  $('#srcTable').innerHTML = sh + '</tbody>';

  $('#glossTable').innerHTML = '<thead><tr><th>Term</th><th>Meaning</th></tr></thead><tbody>'
    + Q.glossary.map(g => '<tr><td class="name">' + esc(g.term) + '</td><td>' + esc(g.meaning)
      + '</td></tr>').join('') + '</tbody>';

  $('#tlTable').innerHTML = '<thead><tr><th>Date</th><th>Event</th><th>Actor</th>'
    + '<th class="num">Days from the start of FY2026</th><th>Source</th></tr></thead><tbody>'
    + Q.timeline.map(t => '<tr><td class="mono">' + esc(t.date) + '</td><td class="name">'
      + esc(t.event) + '</td><td>' + esc(t.actor) + '</td><td class="num">'
      + (has(t.rel) ? N(n(t.rel)) : BLANK) + '</td><td class="dim">' + esc(t.src)
      + '</td></tr>').join('') + '</tbody>';
}

/* ---------- static controls ---------- */
function buildControls() {
  const y = Y(), names = ROWS().map(d => d.name).sort();
  const bl = BASES();
  const blOpts = Object.keys(bl).map(k => ({v: k, l: bl[k].label}));
  const cur = bl[S.baseline[YEAR]] ? S.baseline[YEAR] : y.defBase;
  $('#scenBase').innerHTML = opts(blOpts, cur);
  if (!byName[S.offsetSys]) S.offsetSys = names[0];
  $('#scenOffSys').innerHTML = opts(names, S.offsetSys);
  $('#scenOff').value = S.offset;
  $('#scenSysWrap').style.display = S.offset === 'sys' ? '' : 'none';
  $('#scenGlobal').value = String(S.global * 100);
  $('#scenGlobalV').textContent = Ps(S.global);
  $('#scenText').value = S.text || '';

  if (!byName[S.claim.sys]) S.claim.sys = names.indexOf('Virginia-class Submarine') >= 0
    ? 'Virginia-class Submarine' : names[0];
  if (!bl[S.claim.base]) S.claim.base = y.defBase;
  if (!byName[S.claim.offSys]) S.claim.offSys = names.find(x => x !== S.claim.sys) || names[0];
  /* The comparison line only says something when CRS printed a quantity against
     it, so the default is the largest line that actually has one. */
  if (!byName[S.claim.ref] || unitCost(byName[S.claim.ref]) === null) {
    const priced = ROWS().filter(d => d.name !== S.claim.sys && unitCost(d) !== null)
      .sort((a, b) => n(Y().disc(b)) - n(Y().disc(a)));
    S.claim.ref = priced.length ? priced[0].name : (names.find(x => x !== S.claim.sys) || names[0]);
  }
  const sysOpts = [{v: OWN, l: 'A claim of your own, at any amount'}]
    .concat(names.map(x => ({v: x, l: x})));
  $('#clSys').innerHTML = opts(sysOpts, S.claim.own ? OWN : S.claim.sys);
  $('#clBase').innerHTML = opts(blOpts, S.claim.base);
  $('#clOffSys').innerHTML = opts(names, S.claim.offSys);
  $('#clRef').innerHTML = opts(names, S.claim.ref);
  /* A claim of your own has no baseline to take a percentage of, so the only
     sensible reading of it is an amount. */
  if (S.claim.own) S.claim.type = 'usd';
  $('#clType').value = S.claim.type;
  $('#clType').disabled = !!S.claim.own;
  $('#clOff').value = S.claim.off;
  $('#clOffWrap').style.display = S.claim.off === 'sys' ? '' : 'none';
  $('#clOwnWrap').style.display = S.claim.own ? '' : 'none';
  $('#clOwnLab').value = S.claim.ownLab || '';
  $('#clValLab').textContent = S.claim.type === 'pct' ? 'Percentage' : 'Amount';
  $('#clUnit').style.display = S.claim.type === 'pct' ? 'none' : '';
  $('#clUnit').value = String(S.claim.unit || 1e9);
  $('#clVal').value = S.claim.type === 'pct'
    ? (S.claim.val * 100)
    : S.claim.val;
  $('#clVal').step = S.claim.type === 'pct' ? '0.5' : 'any';

  $('#paySvc').innerHTML = opts(['All services'].concat(PAY.forceServices), S.pay.svc);
  $('#payGrade').innerHTML = opts(GRADES, S.pay.grade);
  $('#payYear').value = String(PAYYEAR());
  $('#payAta').value = String(S.pay.ata * 100);
  $('#payAtaV').textContent = P(S.pay.ata, 1);
  $('#payGradePct').value = String(S.pay.gradePct * 100);
  $('#payGradePctV').textContent = P(S.pay.gradePct, 1);
  $('#payMonth').value = String(S.pay.month);
  $('#payRpa').value = String(S.pay.rpaOn);

  $('#prInfl').value = String(n(S.proj.infl) * 100);
  $('#prInflV').textContent = P(n(S.proj.infl), 1);
  $('#prReal').value = String(n(S.proj.real) * 100);
  $('#prRealV').textContent = P(n(S.proj.real), 1);
  $('#prRepl').value = String(n(S.proj.repl) * 100);
  $('#prReplV').textContent = P(n(S.proj.repl), 0);

  const catOpts = [{v: 'all', l: 'Every category'}].concat(CATS.map(c => ({v: c, l: c})));
  const svcOpts = [{v: 'all', l: 'Every service'}].concat(SVCS.map(c => ({v: c, l: c})));
  $('#dCat').innerHTML = opts(catOpts, S.dview.cat);
  $('#dSvc').innerHTML = opts(svcOpts, S.dview.svc);
  $('#dQ').value = S.dview.q || '';
  $('#cmpCat').innerHTML = opts(catOpts, S.cview.cat);
  $('#cmpPres').value = S.cview.pres;
  $('#cmpQ').value = S.cview.q || '';

  $('#y26').setAttribute('aria-pressed', String(S.year === 2026));
  $('#y27').setAttribute('aria-pressed', String(S.year === 2027));
  $('#yboth').setAttribute('aria-pressed', String(S.year === 'both'));
  $('#footNote').textContent = 'Built ' + DATA.meta.built + ' from ' + DATA.meta.workbook
    + '. Every figure is read from that workbook; none is recalculated here. '
    + 'The workbook itself travels inside this page and can be downloaded from the '
    + 'Workbook button at the top.';
  const amt = $('#clAmt');
  if (amt) amt.textContent = S.claim.type === 'pct' ? ''
    : amountWords(claimBillions()) + '  (' + B(claimBillions(), 4) + ')';
}

const OWN = '__own__';

function setTab(t) {
  S.tab = t;
  $$('.navlink').forEach(b => b.setAttribute('aria-selected', String(b.dataset.tab === t)));
  $$('.tabpanel').forEach(p => p.classList.toggle('on', p.id === 't-' + t));
  /* The masthead belongs to the overview. Repeating it above every section
     would push the working part of each one off the screen. */
  const hero = $('#heroSect');
  if (hero) hero.style.display = (t === 'overview') ? '' : 'none';
  persist();
}

function refresh() {
  reindex();
  buildControls();
  renderOverview();
  renderScenario();
  renderClaim();
  renderPay();
  renderProjection();
  renderCompare();
  renderData();
  renderMethod();
  revealAll();
  persist();
}

function revealAll() {
  $$('.reveal').forEach(el => el.classList.add('in'));
}

/* The workbook is carried in the page as base64 so the reader can take the
   source away with them. Decoding happens only when the button is pressed, so
   it costs nothing until it is wanted. */
const XLSX_B64 = "__XLSX_B64__";
function downloadWorkbook() {
  try {
    const bin = atob(XLSX_B64);
    const buf = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    const blob = new Blob([buf], {
      type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = DATA.meta.workbook;
    document.body.appendChild(a); a.click();
    setTimeout(() => {URL.revokeObjectURL(url); a.remove();}, 1000);
    return true;
  } catch (e) { return false; }
}

function download(name, text, type) {
  try {
    const blob = new Blob([text], {type: type || 'text/plain;charset=utf-8'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = name;
    document.body.appendChild(a); a.click();
    setTimeout(() => {URL.revokeObjectURL(url); a.remove();}, 500);
  } catch (e) {}
}
const numv = (el, div) => {const v = parseFloat(el.value); return isNaN(v) ? 0 : v / (div || 1);};

function setYear(v) {
  S.year = v;
  YEAR = (v === 'both') ? 2027 : v;
  const bl = YEARS[YEAR].baselines;
  if (!bl[S.baseline[YEAR]]) S.baseline[YEAR] = YEARS[YEAR].defBase;
  S.sys = {};
  S.dview.sort = null;
  refresh();
}

function wire() {
  $$('.navlink').forEach(b => b.addEventListener('click', () => setTab(b.dataset.tab)));
  $('#y26').addEventListener('click', () => setYear(2026));
  $('#y27').addEventListener('click', () => setYear(2027));
  $('#yboth').addEventListener('click', () => setYear('both'));
  $('#themeBtn').addEventListener('click', () => {
    const cur = document.documentElement.getAttribute('data-theme');
    const next = cur === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    try {localStorage.setItem('fy2627theme', next);} catch (e) {}
  });
  $('#linkBtn').addEventListener('click', function () {
    persist();
    const t = location.href;
    if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(t);
    this.textContent = 'Copied';
    setTimeout(() => {this.textContent = 'Copy link';}, 1400);
  });
  $('#memoBtn').addEventListener('click', () =>
    download('scenario-memo-fy' + YEAR + '.txt', scenarioMemo()));
  $('#xlsxBtn').addEventListener('click', function () {
    const ok = downloadWorkbook();
    this.textContent = ok ? 'Saved' : 'Not available';
    setTimeout(() => {this.textContent = 'Workbook';}, 1600);
  });

  $('#scenApply').addEventListener('click', () => {S.text = $('#scenText').value; refresh();});
  $('#scenClear').addEventListener('click', () => {
    S.text = ''; S.global = 0; S.sys = {}; S.offset = 'none'; $('#scenText').value = ''; refresh();
  });
  $('#scenCsv').addEventListener('click', () =>
    download('scenario-fy' + YEAR + '.csv', csvScenario(), 'text/csv;charset=utf-8'));
  $('#scenText').addEventListener('change', () => {S.text = $('#scenText').value; refresh();});
  $('#scenBase').addEventListener('change', e => {S.baseline[YEAR] = e.target.value; refresh();});
  $('#scenOff').addEventListener('change', e => {S.offset = e.target.value; refresh();});
  $('#scenOffSys').addEventListener('change', e => {S.offsetSys = e.target.value; refresh();});
  $('#scenGlobal').addEventListener('input', e => {
    S.global = numv(e.target, 100); $('#scenGlobalV').textContent = Ps(S.global); renderScenario();
  });
  $('#scenGlobal').addEventListener('change', () => refresh());

  $('#clSys').addEventListener('change', e => {
    if (e.target.value === OWN) {
      S.claim.own = true;
      if (S.claim.type !== 'usd') {S.claim.type = 'usd'; S.claim.val = 2; S.claim.unit = 1e6;}
    } else {
      S.claim.own = false; S.claim.sys = e.target.value;
    }
    buildControls(); renderClaim(); persist();
  });
  $('#clOwnLab').addEventListener('input', e => {S.claim.ownLab = e.target.value; renderClaim();});
  $('#clOwnLab').addEventListener('change', () => persist());
  $('#clUnit').addEventListener('change', e => {
    S.claim.unit = parseFloat(e.target.value) || 1e9; buildControls(); renderClaim(); persist();
  });
  $('#clBase').addEventListener('change', e => {S.claim.base = e.target.value; renderClaim(); persist();});
  $('#clType').addEventListener('change', e => {
    if (S.claim.own) {e.target.value = 'usd'; return;}
    S.claim.type = e.target.value;
    S.claim.val = S.claim.type === 'pct' ? 0.10 : 1;
    S.claim.unit = 1e9;
    buildControls(); renderClaim(); persist();
  });
  $('#clVal').addEventListener('input', e => {
    S.claim.val = S.claim.type === 'pct' ? numv(e.target, 100) : numv(e.target, 1);
    const amt = $('#clAmt');
    if (amt) amt.textContent = S.claim.type === 'pct' ? ''
      : amountWords(claimBillions()) + '  (' + B(claimBillions(), 4) + ')';
    renderClaim();
  });
  $('#clVal').addEventListener('change', () => persist());
  $('#clOff').addEventListener('change', e => {
    S.claim.off = e.target.value;
    $('#clOffWrap').style.display = S.claim.off === 'sys' ? '' : 'none';
    renderClaim(); persist();
  });
  $('#clOffSys').addEventListener('change', e => {S.claim.offSys = e.target.value; renderClaim(); persist();});
  $('#clRef').addEventListener('change', e => {S.claim.ref = e.target.value; renderClaim(); persist();});

  $('#payYear').addEventListener('change', e => {S.pay.year = +e.target.value; renderPay(); renderClaim(); persist();});
  $('#paySvc').addEventListener('change', e => {S.pay.svc = e.target.value; renderPay(); persist();});
  $('#payGrade').addEventListener('change', e => {S.pay.grade = e.target.value; renderPay(); persist();});
  $('#payAta').addEventListener('input', e => {
    S.pay.ata = numv(e.target, 100); $('#payAtaV').textContent = P(S.pay.ata, 1); renderPay();
  });
  $('#payGradePct').addEventListener('input', e => {
    S.pay.gradePct = numv(e.target, 100); $('#payGradePctV').textContent = P(S.pay.gradePct, 1); renderPay();
  });
  $('#payAta').addEventListener('change', () => persist());
  $('#payGradePct').addEventListener('change', () => persist());
  $('#payMonth').addEventListener('change', e => {
    S.pay.month = Math.max(1, Math.min(12, Math.round(numv(e.target, 1)) || 1));
    e.target.value = String(S.pay.month); renderPay(); persist();
  });
  $('#payRpa').addEventListener('change', e => {S.pay.rpaOn = +e.target.value; renderPay(); persist();});
  $('#payReset').addEventListener('click', () => {S.pay = DEF().pay; buildControls(); renderPay(); persist();});

  [['#prInfl', 'infl', '#prInflV', 1], ['#prReal', 'real', '#prRealV', 1],
   ['#prRepl', 'repl', '#prReplV', 0]].forEach(([sel, key, lab, dp]) => {
    $(sel).addEventListener('input', e => {
      S.proj[key] = numv(e.target, 100);
      $(lab).textContent = P(S.proj[key], dp);
      renderProjection();
    });
    $(sel).addEventListener('change', () => persist());
  });

  $('#dQ').addEventListener('input', e => {S.dview.q = e.target.value; renderData();});
  $('#dQ').addEventListener('change', () => persist());
  $('#dCat').addEventListener('change', e => {S.dview.cat = e.target.value; renderData(); persist();});
  $('#dSvc').addEventListener('change', e => {S.dview.svc = e.target.value; renderData(); persist();});
  $('#dCsv').addEventListener('click', () =>
    download('fy' + YEAR + '-selected-lines.csv', csvData(), 'text/csv;charset=utf-8'));

  $('#cmpQ').addEventListener('input', e => {S.cview.q = e.target.value; renderCompare();});
  $('#cmpQ').addEventListener('change', () => persist());
  $('#cmpCat').addEventListener('change', e => {S.cview.cat = e.target.value; renderCompare(); persist();});
  $('#cmpPres').addEventListener('change', e => {S.cview.pres = e.target.value; renderCompare(); persist();});

  document.addEventListener('click', e => {
    const th0 = e.target.closest ? e.target.closest('th.sortable') : null;
    if (!th0) return;
    const tbl = th0.closest('table');
    const view = tbl && tbl.dataset ? tbl.dataset.view : null;
    if (!view || !S[view]) return;
    const k = th0.dataset.sort;
    if (S[view].sort === k) S[view].dir = -S[view].dir;
    else {S[view].sort = k; S[view].dir = -1;}
    if (view === 'dview') renderData();
    else if (view === 'cview') renderCompare();
    else renderScenario();
    persist();
  });
}

/* ---------- boot ---------- */
/* One named handle on the window, so the page can be inspected in a console and
   exercised by a test harness without loosening anything else. */
window.LAB = {
  DATA: DATA, YEARS: YEARS, get S() {return S;}, set S(v) {S = v;},
  get YEAR() {return YEAR;}, get LEDGER() {return LEDGER;},
  ROWS: ROWS, BASES: BASES, catRows: catRows, reindex: reindex,
  parseLine: parseLine, parseAmount: parseAmount, parseScript: parseScript,
  claimBillions: claimBillions, amountWords: amountWords, downloadWorkbook: downloadWorkbook,
  OWN: OWN,
  matchTarget: matchTarget, compute: compute, payCompute: payCompute,
  payRates: payRates, payRate: payRate, strengthFor: strengthFor,
  projCompute: projCompute, claimCompute: claimCompute,
  scenarioMemo: scenarioMemo, csvData: csvData, csvScenario: csvScenario,
  setYear: setYear, setTab: setTab, refresh: refresh, filtered: filtered,
  version: DATA.meta.version
};

try {
  const th0 = localStorage.getItem('fy2627theme');
  if (th0) document.documentElement.setAttribute('data-theme', th0);
} catch (e) {}
restore();
reindex();
wire();
refresh();
setTab(S.tab || 'overview');
</script>
</body>
</html>
"""

# ---------------------------------------------------------------------------
# Emit
# ---------------------------------------------------------------------------

TITLE = "Defense Budget Lab, FY2026 and FY2027"
DESC = ("Selected weapon systems, personnel pay and committee action for fiscal "
        "years 2026 and 2027, by Alyssa I. Agard, read from a verified workbook "
        "built on CRS R48860 and R49023.")

DASHES = {
    "—": ", ", "–": " to ", "‒": "-", "―": "-",
    "−": "-", "‐": "-", "‑": "-",
}


def de_dash(s):
    """No em dashes and no en dashes anywhere on the page, including in text
    carried out of the workbook."""
    for bad, good in DASHES.items():
        s = s.replace(bad, good)
    return s


ICON_FILES = [
    ("favicon-32x32.png", 'rel="icon" type="image/png" sizes="32x32"'),
    ("favicon-16x16.png", 'rel="icon" type="image/png" sizes="16x16"'),
    ("apple-touch-icon.png", 'rel="apple-touch-icon" sizes="180x180"'),
]


def icon_links(icon_dir):
    """The page must work from a file:// URL with nothing to fetch, so the icons
    travel inside it as data URIs rather than as files beside it. The same files
    are also shipped in the repository, and favicon.ico sits at the repository
    root, where a browser looks for it without being told."""
    if not icon_dir:
        return "", 0
    if not os.path.isdir(icon_dir):
        raise WorkbookMismatch(
            "no icon directory at %s.\n  Pass --icons with the directory that "
            "holds %s, or pass --icons none to build a page without one."
            % (icon_dir, ", ".join(f for f, _ in ICON_FILES)))
    out, total = [], 0
    for name, attrs in ICON_FILES:
        path = os.path.join(icon_dir, name)
        if not os.path.exists(path):
            raise WorkbookMismatch(
                "%s is missing from %s. The icon set needs %s."
                % (name, icon_dir, ", ".join(f for f, _ in ICON_FILES)))
        with open(path, "rb") as fh:
            blob = fh.read()
        total += len(blob)
        out.append('<link %s href="data:image/png;base64,%s">'
                   % (attrs, base64.b64encode(blob).decode("ascii")))
    return "\n".join(out), total


def render(payload, checks, xlsx_path, icons_html):
    data = json.dumps(payload, separators=(",", ":"), allow_nan=False, ensure_ascii=False)
    data = data.replace("</", "<\\/").replace("<!--", "<\\u0021--")
    head = (TEMPLATE_HEAD
            .replace("__TITLE__", TITLE)
            .replace("__DESC__", DESC)
            .replace("__YEAR__", str(datetime.date.today().year))
            .replace("__ICONS__", icons_html))
    body = (TEMPLATE_BODY
            .replace("__BUILT__", payload["meta"]["built"])
            .replace("__WORKBOOK__", payload["meta"]["workbook"]))
    js = (TEMPLATE_JS + TEMPLATE_JS2 + TEMPLATE_JS3 + TEMPLATE_JS4
          + TEMPLATE_JS5 + TEMPLATE_JS6)
    # The workbook travels inside the page, so the reader can take the source away
    # with them and the page stays a single file with nothing to fetch.
    with open(xlsx_path, "rb") as fh:
        blob = base64.b64encode(fh.read()).decode("ascii")
    js = js.replace("__DATA__", data).replace("__XLSX_B64__", blob)
    return de_dash(head + body + js)


# An asset attribute makes the browser fetch something in order to render the
# page. A hyperlink does not: it waits for a click and is the reader's choice.
# The first kind must never point off the file; the second may.
ASSET_ATTR = re.compile(
    r'\b(src|srcset|poster|action|formaction|data|codebase|background|cite|manifest)'
    r'\s*=\s*["\']([^"\']+)["\']', re.I)
LINK_TAG = re.compile(r'<link\b[^>]*\bhref\s*=\s*["\']([^"\']+)["\']', re.I)
REMOTE = ("http:", "https:", "//", "ftp:", "ws:", "wss:")


def selfcheck(html):
    """The page must be self contained and free of em and en dashes."""
    problems = []
    for m in ASSET_ATTR.finditer(html):
        url = m.group(2).strip()
        if url.startswith(REMOTE):
            problems.append("external %s reference: %s" % (m.group(1), url[:90]))
    for m in LINK_TAG.finditer(html):
        url = m.group(1).strip()
        if url.startswith(REMOTE):
            problems.append("external stylesheet or link tag: " + url[:90])
    for tag in ("<iframe", "<object", "<embed", "<frame"):
        if tag in html.lower():
            problems.append("the page embeds %s, which can fetch from the network" % tag)
    for bad in DASHES:
        i = html.find(bad)
        if i >= 0:
            problems.append("dash character %r at offset %d: %r"
                            % (bad, i, html[max(0, i - 50):i + 50]))
    for token in ("__DATA__", "__XLSX_B64__", "__TITLE__", "__DESC__", "__BUILT__",
                  "__WORKBOOK__", "__YEAR__", "__ICONS__"):
        if token in html:
            problems.append("placeholder %s was never substituted" % token)
    return problems


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Build the FY2026 and FY2027 Defense Budget Lab page from the workbook.")
    ap.add_argument("--xlsx", required=True, help="path to the analytics workbook")
    ap.add_argument("--out", required=True, help="path of the HTML file to write")
    ap.add_argument("--open", action="store_true", dest="open_after",
                    help="open the page in the default browser once it is written")
    ap.add_argument("--icons", default=None, metavar="DIR",
                    help="directory holding the favicon PNGs to inline; defaults "
                         "to assets/icons beside this script; pass 'none' to omit")
    args = ap.parse_args(argv)

    if not os.path.exists(args.xlsx):
        raise WorkbookMismatch("no workbook at %s" % args.xlsx)

    if args.icons is None:
        here = os.path.dirname(os.path.abspath(__file__))
        for candidate in (os.path.join(here, "assets", "icons"),
                          os.path.join(os.path.dirname(here), "assets", "icons")):
            if os.path.isdir(candidate):
                args.icons = candidate
                break
    icon_dir = None if (args.icons in (None, "none", "")) else args.icons
    icons_html, icon_bytes = icon_links(icon_dir)

    payload, rd = build_payload(args.xlsx)
    html = render(payload, rd.checks, args.xlsx, icons_html)

    problems = selfcheck(html)
    if problems:
        raise SystemExit("\n  PAGE REJECTED\n  " + "\n  ".join(problems) + "\n")

    outdir = os.path.dirname(os.path.abspath(args.out))
    if outdir and not os.path.isdir(outdir):
        os.makedirs(outdir)
    with open(args.out, "w", encoding="utf-8") as fh:
        fh.write(html)

    sys.stdout.write(
        "wrote %s\n"
        "  %s bytes\n"
        "  %d sheet and header checks passed\n"
        "  %d FY2026 lines, %d FY2027 lines, %d on the joined spine\n"
        "  %d pay grades, active-duty strength %s\n"
        "  %s\n"
        % (args.out, format(os.path.getsize(args.out), ","), rd.checks,
           len(payload["fy26"]), len(payload["fy27"]),
           len(payload["compare"]["rows"]), len(payload["pay"]["grades"]),
           format(int(payload["pay"]["forceSum"]), ","),
           ("%d icons inlined from %s (%s bytes)"
            % (len(ICON_FILES), icon_dir, format(icon_bytes, ",")))
           if icon_dir else "no favicon embedded"))

    if args.open_after:
        import webbrowser
        webbrowser.open("file://" + os.path.abspath(args.out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
