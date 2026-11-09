# FY2026 and FY2027 Defense Budget Lab: build and verification note

Alyssa I. Agard. 22 September 2026. Second revision, with a third and a fourth
recorded in parts 7 and 8 below.

Three files are delivered:

| File | What it is |
| --- | --- |
| `fy2627_budget_lab.py` | the generator, Python 3 plus openpyxl, nothing else |
| `fy2627_budget_lab.html` | the page it emits, 1.0 MB, self contained, with the workbook carried inside it |
| `FY2026_FY2027_Defense_Budget_Analytics.xlsx` | the workbook, repaired (see part 1) |

Run it with:

```
python3 fy2627_budget_lab.py --xlsx FY2026_FY2027_Defense_Budget_Analytics.xlsx \
                             --out  fy2627_budget_lab.html [--open]
```

---

## 1. The workbook had to be repaired first

The page reads the workbook's cached formula results. Before building anything
I audited those results, and the workbook as it stood did not hold up. Two
defects, both invisible until the file was recalculated.

**Twenty-four `#VALUE!` cells and eight silently wrong zeros.** The cached
values in `Personnel_Pay!D68:G72` and `D73:G73` read `#VALUE!`, and
`Year_Compare!B123`, `B124`, `B126`, `B127`, `Pay_Module!B83`, `B84`,
`Pay_Allowances!C72` and `C73` all cached as zero. The consequence on
`Year_Compare` was visible: the bridge residual at `B130` read $85.03B where it
should read zero, so the sheet said its own reconciliation did not close. These
were stale caches, not live errors. Recalculating produced correct figures with
no change to any formula, which is what the delivered workbook now carries.

**Eight cells called a function that does not exist.** `Pay_Module!L69:L76`
each called `_xlfn._LONGTEXT("...","...")`, which is not a function in Excel or
in LibreOffice. The cells carried a stale cached string, so the fault stayed
hidden until recalculation, at which point all eight returned `#NAME?`. Each is
now the plain string concatenation it was meant to be.

**One methodological defect, which mattered more than either.** `Year_Compare`
pulled its figures across with bare references such as `=Systems_Master!$J26`.
A bare reference to an empty cell returns 0. The change columns beside them are
guarded on `=""` and are meant to stay blank whenever a line carries a figure in
only one of the two years. LibreOffice happens to carry emptiness through a bare
reference, so the guards fired there and the fault was invisible; Excel does not.
In Excel, the E-7A Wedgetail and three other lines that R49023 lists without
printing an FY2027 figure would have read as cuts of their entire FY2026
request. Each of the 1,099 pulls on that sheet is now wrapped so the blank
survives the hop between sheets in either engine. That changed 143 displayed
cells from a phantom `$0.00` to a blank and changed nothing else.

The FY2026 sheets were deliberately left alone. `Systems_Master` states its own
convention in its header note, that formulas there read blanks as zero, so
wrapping those pulls would have changed documented behaviour rather than
repaired it.

**After the repair:** 0 formula errors across all 27 sheets. Every difference
against a clean recalculation of the unrepaired original is accounted for: 143
cells on `Year_Compare` that went from a false zero to a blank, and one on
`Pay_Module` that went from `#NAME?` to its text. Every tie-out holds.
`QA_Checks` reports 53 PASS, 3 REVIEW and 21 structural checks all passing;
`QA_FY2027` reports 50 OK and the 1 intended CHECK.

> If you have the workbook open in Excel, close it without saving and reopen it.
> A lock file was present while I worked, and saving the version you had open
> would put the stale caches back.

---

## 2. What the page does

Eight sections, one dataset per fiscal year, switchable between FY2026, FY2027
and both. It works from a `file://` URL with no network of any kind: no content
delivery network, no web font, no charting library, no framework. Every chart is
hand-written inline SVG. Chromium loading the file made **0 network requests**.

The generator verifies **511 sheet and header cells at run time**. No column is
ever read by position without the label beside it being checked first. If a
sheet is missing or a header has moved or been reworded, it names the sheet and
the cell and refuses to write a page.

---

## 3. Verification

### Engine, run in Node against the page's own JavaScript

**124 assertions, 124 passed, 0 script errors.** Among them:

- FY2026: 97 lines; base $114.07B, NDAA $124.86B, appropriation $127.60B, spend
  plan $47.48B, enacted total $175.08B, mandatory dependence 0.2711903. Each
  also re-summed from the individual lines rather than read from the total row.
- FY2027: 104 lines, 101 inside the category tables and 3 in the appendix;
  total request $271.28B, discretionary $203.12B, mandatory $68.16B. All eight
  computed category totals within CRS rounding of the CRS narrative figure, and
  each narrative figure checked against the published list
  `[59.9, 2.1, 18.1, 4.8, 4.9, 15.7, 36.1, 61.5]`.
- Blanks: 0 FY2027 cells populated in the enacted NDAA, enacted appropriation
  and SAC columns; 5 blank discretionary cells; the E-7A Wedgetail carries
  $0.20B for FY2026, a blank for FY2027, and a **blank** change rather than a
  cut to zero.
- Crosswalk: 95 both years, 6 FY2027 only, 2 FY2026 only. Bridge recomputed from
  the spine: $168.56B to $253.59B, change $85.03B, residual 0.
- Pay: E-4 strength 239,169 and Army E-4 95,338; the five services sum exactly
  to the whole-force cost; the FY2026 and FY2027 accrual rates differ as the M-1
  says (0.242282 against 0.201473); O-10 is held at the statutory cap.
- Scenario: add, cut, set, zero, percentage and dollar forms; category and
  service targets; all three offset rules returning the topline where a pool
  exists; a bare number flagged as assumed billions; an unparseable line refused
  rather than guessed.
- Projection: FY2027 gap $48.4296B and cumulative $252.0296B, matching the
  workbook exactly; full replacement closes the gap to zero.

### Page, driven in jsdom

**175 interactions, 174 checks, 0 failures, 0 script errors.** Every one of the
33 panels rendered on load; all eight tabs opened; the composer parsed a
three-line script including a deliberate nonsense line and flagged exactly that
one; all four offset rules; all five baselines; the claim tester across four
offset modes, both change types, all four amount units and a claim of the
reader's own; the pay year switch and all six service scopes; all three
projection dials; every data, compare and scenario column sorted both ways
without losing a row; the year switch in all three positions; the workbook
download; theme toggle; and a shared link restoring the year and the open tab.
The removed phrases, the author's name, the three contact links and the licence
are each asserted individually.

### Fonts, resolved in Chromium

Custom properties are not resolved by jsdom, so the cascade is read in a real
browser: navigation, buttons, table headings, tags and field labels resolve to
`"Segoe UI", Calibri, "Helvetica Neue", Helvetica, Arial, sans-serif`; body
text, headings, running prose and the editorial figures resolve to
`"Times New Roman", Times, serif`. Neither is a web font, so nothing is
fetched.

### Fault injection

Nine deliberate breakages, nine correct and specific refusals: a reworded
header, a removed sheet, a header shifted one column, an altered tie-out, an
FY2027 enacted column filled in, a bridge that stops closing, a deleted line, a
QA check switched to FAIL, and a workbook that was never recalculated. Each
names the sheet and cell at fault and writes no page.

### Independent spot check

A second script re-read the workbook and compared **269 values** sampled from
the page's payload against the cells they came from: 269 matching, 0 not.

### Self-containment

- external asset attributes (`src`, `href`, `srcset`, `action`, and the rest): **0**
- CSS `@import`: **0**; non-`data:` `url()`: **0**
- `fetch`, `XMLHttpRequest`, `WebSocket`, `EventSource`, `sendBeacon`, `Worker`,
  dynamic `import`: **0**
- network requests in Chromium from `file://`: **0**
- em dashes, en dashes, figure dashes, horizontal bars and minus signs anywhere
  in the file: **0**. The only non-ASCII characters are HTML entities.
- one `<script>` tag, one `<style>` tag, everything inline
- no horizontal overflow at 390px
- the embedded workbook decodes **byte identical** to the source file
  (416,920 bytes, same SHA-256) and reopens with all 27 sheets intact

Four `href` values point off the file: your site, your LinkedIn profile, your
email, and the Creative Commons licence deed. All four are ordinary hyperlinks
inside `<a>` elements, which the browser follows only when a reader clicks them.
The self-check now distinguishes those from asset attributes, which make the
browser fetch something in order to render, and still refuses to write a page
that carries any of the latter. The remaining `https://` strings are citation
text inside the sources table.

---

## 4. What the parser could not do

- **No analyst service field exists for FY2027.** `FY2027_Systems` has no
  counterpart to the FY2026 `Service (analyst)` column. The page inherits a
  service label through the crosswalk where a line appears in both years, and
  leaves it blank otherwise rather than guessing. Six FY2027 lines therefore
  read "not classified" in the service column and are excluded from a service
  filter.
- **Basic pay is the January 2026 table in both years.** The workbook carries
  one pay table. Pricing FY2027 switches the burden rates, the caps and the
  composite comparison, but not the underlying basic pay, because no FY2027 pay
  table exists until the raise is set.
- **A category with nothing untouched cannot absorb an offset.** An instruction
  that moves every line in a category leaves the "hold each category constant"
  rule with no pool to draw on, so part of the net stays on the topline. The
  page now says so in words rather than showing an unexplained residual.
- **`Pay_Composite` is sparse by design.** Where a service published no rate for
  a grade, the cell is blank and the spread column reads "not printed". That is
  the source, not a parsing failure.

---

## 5. Where the workbook and the page differ

**Three points where your brief and the workbook disagree, and the workbook won:**

- The brief says `FY2027_Systems` has "109 data rows from row 6" and that 109
  lines were carried. The workbook has **104** data rows, in rows 6 to 109, and
  `QA_FY2027` itself states 104. 109 is the last row number, not a count. The
  page carries 104.
- The brief says `Year_Compare` section B begins "from row 133". Row 133 is the
  section title; the header row labelled "Direction" is row **137**. The reader
  searches for that label between rows 130 and 155 rather than trusting either
  number.
- The brief says `Systems_Master` has 47 columns. It has 47, A through AU. No
  discrepancy; I had noted one earlier and it was my miscount.

**One place the page states something the workbook leaves implicit.** The
crosswalk labels 95 lines "Both years", but only **91** carry a figure in both;
the bridge on `Year_Compare` uses the stricter test and so nets $168.56B rather
than the $170.77B that the looser reading would give. The page originally showed
the looser figure beside the bridge's stricter one, which would have put two
inconsistent numbers on the same screen. It now shows both counts, names the four
lines involved (C-40 Fleet Expansion, E-7A Wedgetail, M10 Booker, and the
Sea-Launched Cruise Missile-Nuclear), and uses the bridge's definition for the
change.

**One disagreement inside the sources, carried through and not reconciled.** The
R49023 summary says the Senate Armed Services Committee proposed authorizing the
amount requested; the committee columns in the same report net to +$0.67B across
27 departing lines. The workbook flags this as its single CHECK and the page
reproduces both figures in the method section. Likewise the three FY2026 REVIEW
rows, where a CRS narrative total disagrees with the sum of CRS's own table
lines: Ground enacted NDAA, munitions reconciliation, and the shipbuilding total
request. Both figures are shown in every case. Nothing is silently reconciled.


---

## 6. Second revision, 22 September 2026

### The claim tester takes any amount, and a claim of your own

The amount field now carries a unit selector beside it: **dollars, thousand,
million or billion**. Typing `2` and choosing million holds the figure as
0.002 billion internally and echoes `$2.00 million ($0.0020B)` under the input,
so the unit you are working in and the unit the model works in are both visible.

The line selector also offers **"A claim of your own, at any amount"**, with a
free-text field for what the money is for. A claim of your own is not a line in
either report, so it has no baseline, no published quantity and no unit cost.
Those rows stay blank and say so; nothing is invented to fill them. What can
still be said is said: the amount in dollars, its share of everything CRS
tabulates, what it would take out of an offset pool, how many units of a line
that does have a published quantity it would buy, what it is worth as an
across-the-board military pay raise, and what it comes to per member of the
active-duty force.

A worked example, exactly as the page renders it: a personal grant of two
million dollars is **0.00098 percent** of the $204.5B FY2027 discretionary
request, worth **0.0021 percent** of a one-year across-the-board pay raise, and
**$2 a member** spread across the force. It buys **0.0001** of a Columbia-class
submarine at the $16.04B a boat CRS implies.

Because a claim of your own has nothing to take a percentage of, the percentage
option is disabled while it is selected and re-enabled the moment you pick a
line again. The comparison line now defaults to the largest line that actually
carries a published quantity, so the "same money in" row says something useful
without being changed first.

### Motion

Easing moved to a softer curve with three named durations. Panels cross-fade
and rise on a tab change; editorial figures stagger in at 60 millisecond
intervals; ledger entries and table bodies fade on rebuild; chart bars grow from
their own baseline, vertical bars from the bottom and horizontal bars from the
left, with axes and labels following; cards lift and their icons scale on hover;
buttons lift and drop a soft shadow; inputs take a focus ring; links underline on
approach. Every one of these is switched off under `prefers-reduced-motion`,
including the animation delays, which the earlier revision did not clear.

### Authorship and licence

You are named in the utility strip, under the masthead, in the page metadata
(`author`, `copyright`, `license`, and the Open Graph article tags) and in a
footer byline carrying your site, your LinkedIn profile and your email. The
footer also carries a **CC BY 4.0** licence statement with the mark drawn as
inline SVG and a link to the deed. The statement separates what is and is not
yours: the underlying CRS reports are United States government works and not
subject to copyright, while the transcription, the crosswalk, the derived
figures and the page itself are your own.

### The workbook travels with the page

The **Workbook** button at the top decodes the spreadsheet out of the page and
saves it. The file is carried as base64 inside the single HTML document, so the
page stays one file with nothing to fetch, and a reader who is handed only the
page still gets the source. Decoding happens when the button is pressed, not on
load, so it costs nothing until it is wanted. The page grew from 455 KB to
1.0 MB to carry it.

### Text changes made as asked

Removed: "511 sheet and header checks passed before this page was written",
", transcribed and checked", and "&middot; 511 header checks passed". The
footer note now says where the workbook came from and that the workbook can be
downloaded, without the check count. Changed: "This page is a reading of two
Congressional Research Service reports" to "This page is a reading of
Congressional Research Service reports". The 511 checks still run on every
build and are still reported on the command line; they are simply no longer
printed on the page.

### Fonts

Navigation, buttons, the year switch, table headings, field labels, tags and
the utility strip use Segoe UI with Calibri behind it. Everything a reader
actually reads, including headings, running prose, the editorial figures and
table cells, stays in Times New Roman. Figures stay monospaced.


---

## 7. Third revision: prepared for publication

The workbook was written for private use and said so in places a reader outside
that working folder could not check. Three changes make it publishable as it
stands.

**The reference list is now a reference list.** `Sources_FY2027` carried a
column headed "File in the reference folder", whose values were the PDF
filenames as they sat on one machine, and it closed with four notes about that
folder: which files were byte-identical duplicates, which had no usable text
layer, and which had a misleading name. None of that means anything to anyone
else, and a public document should not be pointing at a private directory. The
sheet has been rebuilt in the same six columns as the FY2026 `Sources` sheet,
with the reference numbering continued rather than restarted, so the two sheets
now form one list running from S1 to S17. Each entry carries its title,
publisher, date, what it supplies and a link. The page presents the two sheets
as a single table, and the generator refuses to build if the two sheets ever
reuse a reference number.

Three further cells named internal files and have been rewritten to say the same
thing in checkable terms. `Pay_Composite` sections A and B pointed at the source
PDFs by filename and now cite S15. The quality assurance finding about the
folder's duplicate files now states the substantive point, which is that every
source document was checked for integrity and a usable text layer before any
figure was taken from it, that one document had almost no text layer and was not
used, and that duplicate copies were identified and read once. Nothing that was
verified has been dropped; only the private paths are gone. A scan of every
string cell in all 27 sheets returns no remaining reference to a folder, a local
filename, or the working set. The only remaining occurrences of `.pdf` anywhere
are inside published citation URLs.

**The workbook carries its own authorship and licence.** The Guide sheet gained a
block naming the author, the three contact points, the CC BY 4.0 licence with a
link to its text, and the citation line. The same information is written into the
file's own document properties, so it survives being renamed, emailed or
detached from this repository. The Guide's contents list had stopped at the
FY2026 sheets; the eight sheets added for FY2027 and for the pay work are now
listed with the rest, and the Guide's title and standfirst describe both years.

**Everything downstream was rebuilt.** Header verification rose from 511 checks
to 513 with the new source columns. All 124 engine assertions, all 174
interaction checks and the Chromium pass repeat clean, with zero script errors
and zero network requests. The copy of the workbook carried inside the page is
byte identical to the one in `data/`, and the page's source table now shows all
seventeen references with working links.

## 8. Fourth revision, 1 November 2026

**A statement about the compensation tables was wrong, and has been corrected
everywhere it was published.** The workbook, the page, this repository's README
and the report all said, in different words, that the Selected Military
Compensation Tables had not been published since the edition of 1 January 2023.
Two things show otherwise. A CY2024 edition is publicly posted on the
Department's military pay site, and it is the most recent edition that is. And
CRS IF10532, *Defense Primer: Regular Military Compensation*, updated 7 July
2026, sources its Table 1 to the Selected Military Compensation Tables of
1 January 2026, page B-3, "on file with CRS": that edition exists but is not
publicly posted. The corrected statement is in two parts accordingly. The cells
changed are `Sources_FY2027` S13, `Pay_Allowances` A3, A70, A87 and A95,
`QA_FY2027` A71 and B71, `Pay_Module` A102, which still said the tables had not
been supplied, and the Guide's line for `Pay_Allowances`; on the page, the lead
paragraph of the regular military compensation table.

**The source list runs to S18.** S14, IF10532, had been dated only "Updated
2026". It now names the version of 7 July 2026, the one held in the reference
set, and says what its Table 1 supplies. S18, the CY2024 edition of the
compensation tables, is new; it is cited for the fact that it is the most recent
edition publicly posted and supplies no figure. To keep a blank row between the
list and the notes beneath it, which is where the generator stops reading, the
notes moved down one row.

**The scaling behind regular military compensation has been tested against a
published figure.** `Pay_Allowances` section H sets the model, at both priced
years, against the six grades in IF10532 Table 1, every comparison a live
formula. The model does not reproduce them. Total compensation lands above the
published figure at all six grades, by about 1 to 3 percent for the officer
grades and 5 to 6 percent for the enlisted ones, but the components disagree.
Published housing grew by 13 to 15 percent between the 2023 and 2026 editions at
five of the six grades, which supports carrying the 2023 shape forward with one
factor per group, but the factors fitted to the M-1 run housing 14 to 24 percent
above the published allowance everywhere except E-1. The M-1 officer subsistence
line pays for about 4 percent more officer-years than the 30 September 2025
strength in FY2026, and about 8 percent more in FY2027, so the factors carry a
headcount difference as well as a rate. The marginal rate behind the tax
advantage agrees with the published one at E-1, E-8 and O-6, and the tax
advantage runs about a quarter low at E-5, O-1 and O-4. Nothing upstream reads
section H, and the burden rates derived from the M-1
are ratios of M-1 dollars, so no figure on the page moved.

**The documented build now reproduces the committed page.** The page committed
in September carried favicons with the same pixels as the files in
`assets/icons` but different bytes. Rebuilt with the command in the README,
which reads `assets/icons`, the page would therefore have differed from the
committed copy, and the drift check in `ci-workflows/build.yml` would have failed
on its first run. The page is now built from `assets/icons` as the command says.
The icons look the same, and they account for about 23 kilobytes of the page's
growth; the larger workbook it carries accounts for most of the rest.

**What was checked, and how.** The workbook was edited with openpyxl and
recalculated in LibreOffice 24.2.7.2, the version that last saved it, with zero
formula errors across 15,635 formulas. A cell by cell comparison with the
September workbook, covering formulas, cached values and cell styles on every
sheet, finds no change outside the cells listed above and the new section H.
The generator verifies 513 sheet and header cells, as before. The page's data
payload was compared with the September build field by field: every figure is
unchanged, and the differences are the corrected text, the three source entries
and the build date. `tools/check_page.py` passes, the workbook carried inside the
page is byte identical to the one in `data/`, and Chromium loads the page with
no script errors and no network requests. The Node engine and jsdom harnesses
described in part 3 were not re-run for this revision; the page's script changed
only in one paragraph of text, and its figures did not change.

**One addition before publication, 9 November 2026.** Source entry S12 in
`Sources_FY2027` now names the bills and committee reports behind R49023's
committee columns: H.R. 8800 (H.Rept. 119-698), S. 4784 (S.Rept. 119-127) and
H.R. 9495 (H.Rept. 119-715). It records that the House Appropriations Committee
approved H.R. 9495 on 24 June 2026, by 34 to 27, while R49023 dates the
committee's report 26 June 2026, so that the two dates are not read as a
conflict, and that H.Rept. 119-715 carried no funding table quantities. On the
page, the FY2027 overview gains one sentence stating that the fiscal year began
on 1 October 2026 and that the Department is funded by continuing resolution at
the FY2026 rate until Congress acts. The workbook change was made directly in
the file, to that one cell's text and its row height, so no formula was
rewritten and nothing needed recalculating; a cell by cell comparison with the 1
November workbook, covering formulas, cached values and cell styles on every
sheet, finds that cell and nothing else. The 1 November page rebuilds byte for
byte from its own workbook once the build date is normalised, so the build is
reproducible, and the new page differs from it only in the source entry, the
overview sentence, the workbook it carries and the build date. The generator
verifies 513 sheet and header cells, `tools/check_page.py` passes, the workbook
carried inside the page is byte identical to the one in `data/`, now 429,376
bytes, and Chromium loads the page with no script errors and no network
requests.
