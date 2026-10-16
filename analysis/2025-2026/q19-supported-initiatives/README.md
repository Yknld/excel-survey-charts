# Q19 supported and funded initiatives — provisional extraction

## Report-ready workbook

[Download the combined Excel workbook](Q19_report_examples_2025-2026.xlsx). Its first sheet contains 50 concrete examples selected for a 5–10 page report table. Each row brings together the department, PCH unit where applicable, program, project, recipient, funder, amount, period, communities, activities, participants, reported results, a paste-ready summary and the complete source excerpt. The workbook also retains all 267 initiative mentions and a filtered review list.

Processed 320 original answer entries; 167 contain extracted support records.
Extracted 267 initiative mentions and 30 reported monetary facts.
API cost (uncached input pricing): US$1.3931.
Evidence checks needing correction: 0. Review notes/issues: 297.

## How to use
Initiatives.csv: one reported initiative mention per answer, not unique projects.
Financial_facts.csv: linked monetary statements. Amounts deliberately remain as written, with
currency, qualifier, scope, status and period. NO total is calculated: parent budgets, project
awards, shared reporting and multiple years can overlap. Blank means not stated.
Source_coverage.csv: every original Q19 answer preserved unchanged, including answers with no
extracted initiative. Review_issues.csv: uncertainty and evidence checks for human review.
Duplicate_candidates.csv: matching normalized names; similarity is not proof of duplication.

## Scope and limits
Original survey and PCH consolidated answers are tagged separately. PCH units are subdivisions
of one institution, not additional federal institutions. Ambiguous support is not classified
as money. Extraction is AI-assisted and provisional; exact quote checks verify textual presence,
not all interpretations or completeness. No cross-source merging and no original source edits.
The fixed extraction schema and prompt are in extract_q19_support.py. Cached raw API responses,
source snapshot and usage ledger make the run reproducible. No financial values inferred from
project counts. No currency inferred from a bare dollar sign. No public upload performed.

## Source-based review
21 field corrections logged in Review_corrections.json. Funding_terms.csv retains percentage conditions and eligibility thresholds separately. Exports use extraction_reviewed.json; raw model output is retained.
