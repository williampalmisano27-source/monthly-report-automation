# Verification record

Verified locally on September 19, 2026, on macOS using Python 3.14.4, openpyxl 3.1.5 and pytest 9.1.1.

| Check | Actual result |
|---|---|
| Fresh virtual environment, `python -m pip install -e '.[dev]'` | Installed successfully with declared dependencies |
| `python -m pytest -q` | 71 tests passed |
| `python -m report_rollover.samples --output demo` | Generated 12 valid reports and 3 invalid examples |
| `report-rollover --input demo/input --output demo/output` | Created 12 next-month reports and the audit CSV |
| Independent inspection of committed examples and fresh demo | All 96 opening balances and audit rows matched the expected quantities in each batch |
| CLI run against invalid examples | Exit 2, three file/cell-specific messages, no published output batch |
| Source preservation and injected save failure | Automated tests confirm unchanged source bytes and no partial publication |

The tests were developed before their corresponding implementation and observed failing before passing. Formatting comparisons use underlying openpyxl style values, because comparing its style proxy objects directly reports inequality even for the same cell.

The environment was freshly created, but only macOS / Python 3.14 was exercised. The declared Python 3.11 minimum and Windows instructions have not been independently tested. Native Microsoft Excel and Google Sheets recalculation/rendering have not been verified. Formula values are independently checked in Python; previews may calculate formulas using a separate rendering engine.

## Independent review and fixes

A separate read-only reviewer examined the whole initial implementation and independently reproduced the ordinary demo. Three important findings were reproduced with failing regression tests and fixed:

- Excel output could lose precision or become blank for very large balances. Quantities and calculated balances now must survive a 15-significant-digit check, and saved workbooks are reopened to verify their opening balances agree with the decimal calculation before publication.
- Corrupt style indices, shared-string indices, and font definitions could leak library exceptions. These now become filename-specific validation errors; reader and actual CLI regression tests cover all three cases.
- Literal text resembling a formula could pass. A closing cell must now have both the expected expression and the formula cell type.

The resulting complete suite passes 71 tests. This is a regression-verified fix pass; no second independent review was performed.

## Decisions and remaining limits

- Numerical support is explicitly limited to finite, 15-significant-digit values that survive serialization. This favors correct balances over accepting unusually large inputs; some otherwise numeric spreadsheets are rejected.
- Unsupported charts, macros, embedded objects and external links remain outside the template guarantee. They may not survive workbook round-tripping, so use plain supported templates.
- Publication assumes no concurrent source edits or destination manipulation. Handled failures clean up staging; abrupt process termination can leave a temporary directory, and no cross-process locking is provided.
- Native Excel/Google Sheets behavior was not tested. PNG previews were rendered from the example files using a separate spreadsheet engine and inspected for legibility; actual application rendering and recalculation may differ.
- Python 3.11 and Windows remain unverified. The fresh-install check ran on macOS / Python 3.14.4; other environments may have compatibility issues.

Deferred minor: a valid uppercase `.XLSX` filename is processed, but the CLI's success-message count omits it. Output files and the audit CSV are correct; use lowercase `.xlsx` names for an accurate displayed count.
