# Case study: carrying monthly material balances forward

## The task

A recurring spreadsheet workflow closes each month with a material balance, then begins the next month with that same quantity. Across multiple locations, manual copying can introduce missing records, stale receipts, or a report prepared for the wrong month.

This independent portfolio project demonstrates that workflow using 12 synthetic location reports with eight materials each. It was inspired by a public buyer request, but no client commissioned or accepted this implementation. Development used AI assistance.

## The deliverable

A local Python command reads a documented Excel format and produces the following month's reports plus an audit CSV. It stops before publishing an output batch if any input fails validation. The original files remain unchanged.

For location LOC-01, material MAT-001 has September opening stock of 111 kg, receipts of 21 kg, and usage of 6 kg. Its closing balance is 126 kg. The October copy begins at 126 kg, resets receipts and usage to zero, and retains the closing formula. This exact example is included in the repository.

## Engineering choices

**A specific template keeps the promise testable.** The program checks the expected columns and formula. It rejects unknown formulas instead of guessing what a customer's workbook means.

**Decimal arithmetic avoids dependence on saved formula caches.** Some generated spreadsheets contain formulas without cached results. The application computes the supported expression directly, while leaving the actual formula in the output for later spreadsheet use.

**Batch validation and staging prevent partial results after handled failures.** All inputs pass validation first. Files are then written to a temporary sibling directory and published together. An injected disk-write failure verifies that no incomplete output directory is published.

**An audit file makes review possible without Excel.** Each row names its source, location, material, old and new month, previous closing and new opening. Text with formula-like prefixes is escaped for safer spreadsheet import.

## Evidence

The automated suite checks all 96 example records against independently derived expected quantities. It also checks formula preservation, representative styling and print settings, source-file hashes, malformed workbooks, validation errors, calendar boundaries, repeat-run refusal and staging cleanup.

The [verification record](verification.md) lists the tested environment and commands. These results demonstrate correctness within the defined template. They do not establish compatibility with a specific customer's spreadsheet, native Excel rendering, measured customer time savings, or demand for a standalone software product.

## Adapting this to paid work

The next inputs needed from a buyer are anonymized source and expected-output examples, the list of formulas and formatting that must survive, material identifier rules, and late-correction rules. Those samples determine whether the template can be mapped simply or needs a different implementation and quote.
