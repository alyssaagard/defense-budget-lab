<p align="center">
  <img src="assets/icons/android-chrome-192x192.png" width="96" height="96" alt="">
</p>

<h1 align="center">Defense Budget Lab, FY2026 and FY2027</h1>

<p align="center">
  Selected weapon systems, personnel pay, and what a change to one would cost.<br>
  <a href="https://alyssaagard.dev">Alyssa I. Agard</a>
  &middot; <a href="https://creativecommons.org/licenses/by/4.0/">CC BY 4.0</a>
</p>

---

An interactive reading of two Congressional Research Service reports on funding
for selected weapon systems, joined line by line across fiscal years 2026 and
2027 and set beside a model of what military pay actually costs. The page opens
from a file, runs entirely offline, and makes no network request of any kind.
Everything it shows is read from the workbook in `data/`, and the workbook
travels inside the page, so a reader who is handed only the HTML file still has
the source.

**[Open the page](https://alyssaagard.github.io/defense-budget-lab/)** ·
**[Read the verification note](docs/VERIFICATION_NOTE.md)**

## What is here

| Path | What it is |
| --- | --- |
| `index.html` | The page. One file, 1.0 MB, no dependencies, works from `file://`. |
| `src/fy2627_budget_lab.py` | The generator. Python 3 plus openpyxl, nothing else. |
| `data/FY2026_FY2027_Defense_Budget_Analytics.xlsx` | The workbook. 27 sheets, 15,352 live formulas, the single source of truth. It carries its own authorship, licence and full reference list. |
| `docs/VERIFICATION_NOTE.md` | What was checked, what the parser could not do, and where the sources disagree with themselves. |
| `tools/check_page.py` | Confirms a built page fetches nothing and carries its own workbook. |
| `ci-workflows/` | Two GitHub Actions workflows, to be moved into `.github/workflows/`. See the README there. |
| `assets/icons/` | The icon set. |

## Publishing

The page is committed as `index.html` at the repository root, so GitHub Pages
serves it with no build step. Turn it on under **Settings, Pages** and set the
source to **GitHub Actions**; the workflow in `.github/workflows/pages.yml`
publishes the repository as it stands on every push to `main`.

The two workflow files live in `ci-workflows/` rather than `.github/workflows/`,
and `ci-workflows/README.md` says why and gives the one command that moves them.
Read them before you do: a workflow is code GitHub runs on your account.

> The two URLs in this README and in `CITATION.cff` assume the repository sits
> at `alyssaagard/defense-budget-lab`. If you name it something else, change
> them in both files and nowhere else.

## Rebuilding the page

```bash
pip install -r requirements.txt
python src/fy2627_budget_lab.py \
  --xlsx data/FY2026_FY2027_Defense_Budget_Analytics.xlsx \
  --out  index.html
```

Add `--open` to open the result in a browser. The generator verifies 511 sheet
and header cells before it writes anything; if a sheet is missing or a column
heading has moved, it names the sheet and the cell and refuses to build. That
is deliberate. A page that quietly reads the wrong column is worse than no page
at all.

Because `openpyxl` reads cached formula results and cannot evaluate formulas
itself, the workbook has to carry the results of its own last recalculation. If
you edit it with a script, open it in Excel or LibreOffice, let it recalculate,
and save before rebuilding. The generator will tell you if you forget.

## How this was built

I began the way anyone begins with a Congressional Research Service report,
which is to say by reading it. The FY2026 report, R48860, prints its selected
weapon systems across seven category tables and an appendix, and the FY2027
report, R49023, does the same for a different and overlapping set of programmes.
My first pass was entirely manual: I read the tables, wrote the figures down,
and built the FY2026 spine by hand so that I would know what the data actually
looked like before I asked a machine to touch it. That manual pass turned out to
be the most valuable hour of the project, because it is where I learned the
thing that governs everything else. CRS leaves cells genuinely blank. A
programme that the House authorised but the appropriators never took up simply
has nothing printed in that column, and a programme that appears in one year's
report and not the next has nothing printed at all. A blank is not a zero, and
the difference between the two is the difference between reporting that a
committee cut a programme and reporting that nobody wrote a number down.

Doing the FY2027 report the same way was not realistic. It carries 104 lines
across eight categories with up to eight funding columns each, and hand
transcription at that volume is a reliable way to introduce errors you will
never find. So I moved to Python, and the first thing I tried did not work. The
obvious approach is to pull the text out of the PDF and split it on whitespace,
but a flowed text dump collapses those meaningful blanks, and the moment a row
is missing a figure every value after it slides one column to the left. Whole
programmes ended up attributed to the wrong chamber. What works instead is
extracting each word with its coordinates using `pdfplumber`, taking the right
edge of each printed figure, and assigning it to whichever column header sits at
or to the left of that edge. Right-aligned numbers under a fixed header row make
this reliable, and a blank stays blank because nothing lands in that column's
span. Getting the anchors right took some care of its own: the header marker in
one table merged with its superscript footnote letter until I split words on
font size as well as position, and the section headings inside the tables were
being read as programme rows until I learned to test only the first token's
font, because CRS appends an italic product suffix to those lines.

Extraction that looks right is not the same as extraction that is right, so
nothing was built on it until it had been checked against a second, independent
source: me. I transcribed the tables a second time by hand into a Python
structure, without looking at what the parser had produced, and machine-compared
the two. That comparison covered 607 cell slots and came back with zero
mismatches, and the rule I set for myself was that the workbook would not be
started until that number was zero. Two of the corrections it surfaced were
mine, not the parser's.

Building the workbook was `openpyxl`, roughly fifteen thousand live formulas
written programmatically across twenty-seven sheets. Every computed cell is a
formula rather than a pasted Python result, which costs more to write but means
the arithmetic is auditable in the file itself and the scenario levers actually
recalculate when you change them. The workbook follows a consistent convention:
blue for source data as published, black for formulas, green for links between
sheets, and a pale fill for the cells you are meant to edit. A quality assurance
sheet sets every computed category total beside the figure CRS states in its own
narrative and reports the gap, and where the two genuinely disagree, which
happens in four places across the two reports, both figures are kept and the
disagreement is named. Nothing is quietly reconciled.

The pay module was the part that would not come from a table. Regular military
compensation is basic pay plus the housing and subsistence allowances plus the
federal income tax advantage that follows from those allowances not being taxed,
and the Department publishes the resulting figures in its Selected Military
Compensation Tables. The most recent edition publicly posted is CY2024, a
1 January 2026 edition exists but is not publicly posted, and an FY2027 edition
cannot exist until the FY2027 pay raise is set, so there was nothing current to
look up. I rebuilt it instead: the published 2023 tables, the edition in hand
when the model was built, supply the
structure, meaning the share of each grade drawing cash housing and the shape of
the allowance across grades, and the M-1 budget exhibit for the priced year
supplies the level, with one scale factor per group fitted so that the grades
sum exactly to the housing dollars the M-1 itself reports. Both housing
reconciliations close to zero. The tax advantage is grossed up correctly, at
`allowances × t / (1 - t)` rather than the `allowances × t` that a first pass
tempts you into, and the marginal rate is blended from the published married
share across the single and married-filing-jointly brackets. The result is
labelled for what it is, housing per member drawing cash rather than a published
allowance rate, because the M-1 line it is fitted to includes the overseas
housing allowance.

That construction has now been tested against the six grades for which CRS
publishes the 1 January 2026 figures (IF10532, Table 1), and it does not
reproduce them. The total lands within about six percent of the published
figure at every grade, but housing and subsistence run above the published
allowances at every grade, and the tax advantage runs about a quarter low at
E-5, O-1 and O-4, most likely because the marginal rate is taken at basic pay
alone, before the allowances carry taxable income into the next bracket.
Section H of `Pay_Allowances` sets out the comparison.

The page came last, and it is generated rather than written. The generator reads
the workbook and lays it out; it never recomputes a figure the workbook already
carries, so there is exactly one place where any number can be wrong. Every
sheet name and every column heading it depends on is checked at run time before
a byte is written. Charts are hand-written inline SVG, because a charting
library would mean a network request and this page is meant to survive being
emailed to somebody and opened on a plane.

## Verification

The page ships with its checks. Before each release the generator's JavaScript
is extracted and run in Node against every tie-out the project rests on, the
whole page is driven in jsdom, and the result is opened in Chromium to confirm
it fetches nothing. The current numbers:

| Check | Result |
| --- | --- |
| Sheet and header cells verified at build time | 513 |
| Engine assertions in Node | 124 passed, 0 failed |
| Interactions driven in jsdom | 175, across 174 checks, 0 failed |
| Script errors | 0 |
| Fault injections, each naming its own cause | 9 of 9 |
| Values spot-checked back against workbook cells | 269 of 269 matching |
| Network requests from `file://` in Chromium | 0 |
| External asset references, em dashes, en dashes | 0 |

`docs/VERIFICATION_NOTE.md` sets out what each of those covers, along with the
three defects found in the workbook and repaired before the page was built, the
four things the parser could not do, and the places where the sources contradict
themselves.

## Sources

Eighteen references, numbered S1 to S18, are listed in full on the workbook's
two `Sources` sheets and reproduced on the page under **Method and sources**,
each with its publisher, its date, what it supplies and a link. The principal
ones are:

- Congressional Research Service, **R48860**, *FY2026 Defense Budget: Funding for
  Selected Weapon Systems*, updated 2 July 2026.
- Congressional Research Service, **R49023**, *FY2027 Defense Budget: Funding for
  Selected Weapon Systems*, version 4, 3 September 2026.
- Office of the Under Secretary of Defense (Comptroller), **M-1** military
  personnel programs, FY2026 and FY2027 President's Budget.
- Office of the Under Secretary of Defense (Comptroller), **Military Composite
  Standard Pay and Reimbursement Rates**, Tab K, FY2026 and FY2027 editions.
- Office of the Under Secretary of Defense for Personnel and Readiness,
  **Selected Military Compensation Tables**, 1 January 2023, with the CY2024
  edition cited as the most recent publicly posted.
- Congressional Research Service, **IF10532**, *Defense Primer: Regular Military
  Compensation*, updated 7 July 2026, for the definition of regular military
  compensation and the published figures the pay model is checked against.
- Executive Order of 18 December 2025, adjusting certain rates of pay, effective
  1 January 2026.
- Congressional Research Service **IF10684** and **IF10685**, for active-duty
  strength by pay grade from Defense Manpower Data Center counts as of
  30 September 2025.

Reports of the Congressional Research Service and budget exhibits of the
Department of Defense are works of the United States government and are not
subject to copyright.

## Licence

Released under the [Creative Commons Attribution 4.0 International
Licence](https://creativecommons.org/licenses/by/4.0/). You may share and adapt
this work, including commercially, so long as you give credit, link to the
licence, and say whether you changed anything.

The underlying Congressional Research Service reports and Department of Defense
budget exhibits are works of the United States government and are not subject to
copyright. The transcription, the crosswalk between the two fiscal years, the
derived figures, the workbook and this page are my own.

If you use this, please cite it. `CITATION.cff` has the metadata, and GitHub
will format it for you from the sidebar.

## Contact

Alyssa I. Agard
[alyssaagard.dev](https://alyssaagard.dev) ·
[linkedin.com/in/alyssaagard](https://www.linkedin.com/in/alyssaagard/) ·
[economics@alyssaagard.dev](mailto:economics@alyssaagard.dev)

This is not a government publication and carries no endorsement.
