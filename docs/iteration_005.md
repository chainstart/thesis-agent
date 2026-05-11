# Iteration 005 - Generalized Acceptance Pass

## Goal

Move the system away from per-paper fixes and toward reusable thesis processing rules. The acceptance criterion for this iteration is that all bundled draft samples can be processed by the same `batch-process` workflow and pass the quality gate.

## Changes

- Added robust TOC normalization:
  - handles TOC headings inside Word content controls;
  - separates visible TOC headings from TOC field content when needed;
  - prevents TOC headings from carrying section breaks that push entries to the next page.
- Added section cleanup:
  - removes empty paragraphs immediately after section breaks;
  - merges empty section-only paragraphs into the previous content paragraph to avoid blank front-matter pages.
- Improved visual checks:
  - detects caption anchors using PDF text coordinates plus rendered PNG ink around the caption;
  - avoids flagging thin flowcharts and normal chapter-ending sparse pages as blockers.
- Improved content review:
  - audits DOCX structural text before falling back to PDF text;
  - counts unnumbered GB/T-style reference lines;
  - recognizes loose citation forms such as `[6,]`.
  - ignores tabbed TOC entries when building chapter bodies;
  - finds test/debug chapters by title semantics instead of fixed chapter number.
- Added regression coverage for unnumbered references and loose citations.

## Verification

Commands:

```bash
PYTHONPATH=src python3 -m pytest -q
PYTHONPATH=src python3 -m thesis_agent batch-process \
  --template /path/to/private-format-template.doc \
  --inputs /path/to/private-drafts \
  --out runs/acceptance
```

Results:

| Draft | Gate | Score | Blockers |
| --- | --- | ---: | --- |
| 私有样本 A | PASS | 96 | 0 |
| 私有样本 B | PASS | 96 | 0 |
| 私有样本 C | PASS | 96 | 0 |
| 私有样本 D | PASS | 96 | 0 |
| 私有样本 E | PASS | 96 | 0 |

The full generated outputs are under `runs/acceptance/`. Each processed thesis includes `final.docx`, the audit report, content improvement plan, process report, and VLM-ready visual review package.

## Remaining Work

- The content plan is currently advisory. Future iterations should optionally write controlled content improvements into the DOCX, especially for weak test-method sections.
- Reference formatting is counted more robustly, but automatic GB/T 7714 normalization has not yet been implemented.
- The visual package is generated for external VLM review; the feedback is not yet automatically parsed back into another fix pass.
