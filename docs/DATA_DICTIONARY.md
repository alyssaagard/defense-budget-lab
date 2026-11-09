# Workbook sheet guide

`data/FY2026_FY2027_Defense_Budget_Analytics.xlsx` carries 27 sheets. This is
what each one holds and which of them the page reads.

## Conventions used throughout

- **Blue** is source data as published, and is the only thing you should edit.
- **Black** is a formula.
- **Green** is a link pulled from another sheet.
- A **pale yellow fill** marks a scenario lever, meant to be changed.
- A **blank cell means CRS printed no figure on that line.** It is not a zero
  and it is not a cut. The FY2026 sheets state their own narrower convention in
  their header notes, where formulas read a blank as zero for summation; the
  FY2027 sheets and the crosswalk keep blanks blank all the way through.
- Currency is `$#,##0.00`, zeros show as a hyphen, negatives sit in parentheses.
- Percentages are stored as fractions.

## The sheets

| Sheet | What it holds | Read by the page |
| --- | --- | --- |
| `Guide` | How to use the workbook. | no |
| `Dashboard` | A one screen summary of FY2026. | no |
| `Systems_Master` | The 97 FY2026 lines from R48860, every column CRS printed, plus derived metrics. Rows 6 to 102. | yes |
| `Category_Summary` | FY2026 rollups by category, by service and by subcategory. | yes |
| `Congress_Actions` | What each chamber did to each FY2026 line. | no |
| `Reconciliation` | P.L. 119-21 request against the DOD spend plan, including Table A-1 line by line. | yes |
| `Unit_Costs` | Budget line cost per unit for every FY2026 line carrying a quantity. | no |
| `Scenario_Builder` | The FY2026 scenario levers. | no |
| `Claim_Tester` | The FY2026 proportion checks. | no |
| `Personnel_Pay` | Grade level cost model for a FY2026 raise, including cost by service. | yes |
| `Projection` | The reconciliation cliff, FY2027 to FY2031. | yes |
| `Topline` | DOD toplines, reconciliation, the supplemental, and the selected systems in context. | yes |
| `Timeline` | Dated legislative actions on the FY2026 request. | yes |
| `Pay_Table_2026` | The January 2026 monthly basic pay table, 24 grades by 22 years of service steps. | yes |
| `Force_Structure` | Active duty strength by pay grade and service. 1,285,084 in total. | yes |
| `QA_Checks` | FY2026 tie-outs against the CRS narrative, plus structural checks. | yes |
| `Definitions` | Glossary and field dictionary. | yes |
| `Sources` | The FY2026 source list. | yes |
| `Lists` | Dropdown sources. | no |
| `FY2027_Systems` | The 104 FY2027 lines from R49023, with provenance columns carrying the CRS page and footnote for each. Rows 6 to 109. | yes |
| `FY2027_Categories` | FY2027 category totals, computed and as CRS states them, with the gap. Plus Appendix Table A-1 and the investment accounts. | yes |
| `Year_Compare` | The crosswalk. 103 rows joining the two years, the bridge that reconciles the totals with the change column, and the lines present in one year only. | yes |
| `Pay_Composite` | Composite standard pay rates per grade per service, FY2026 and FY2027 editions. | yes |
| `Pay_Module` | The two year pay model: M-1 derived rates, assumptions, and the 24 grade cost table. | yes |
| `Pay_Allowances` | Allowances and regular military compensation per grade, with the housing reconciliation and, in section H, a cross-check against the published 1 January 2026 figures for six grades (CRS IF10532, Table 1). | yes |
| `QA_FY2027` | Every check behind the FY2027 extension, in eight lettered sections. | yes |
| `Sources_FY2027` | The FY2027 source list. | yes |

## Things that are blank on purpose

- Every FY2027 enacted authorization, enacted appropriation and Senate
  appropriations column. Nothing had been enacted and the Senate appropriators
  had not reported. The generator asserts these are empty and refuses to build
  if they are not, which is correct until FY2027 action actually happens.
- Five FY2027 discretionary cells, on lines R49023 lists without printing a
  figure.
- The change columns on `Year_Compare` wherever a line carries a figure in only
  one of the two years.
- `Pay_Composite` cells for grades a service did not publish a rate for.
