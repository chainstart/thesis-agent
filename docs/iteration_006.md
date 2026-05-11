# Iteration 006 - Hard Gates and Iterative TOC Repair

## Goal

Address failures found by visual review on the Yang draft and make those fixes reusable across all bundled samples.

## Changes

- Added hard gates for:
  - Arabic page numbers before the main body;
  - stale TOC page labels after rendering;
  - missing acknowledgements through the expected-section baseline.
- Added template red-text checklist extraction:
  - writes `template_red_checklist.md/json` for each `process` run;
  - preserves the source red-text requirements as checklist items.
- Added a DOCX structure hard gate for the red-text main-heading rule:
  - 一级标题小二号黑体居中;
  - 段前 0 磅，段后 12 磅;
  - 每一章另起页.
- Reworked section handling so front matter stays Roman and the first body section starts at Arabic page 1.
- Added controlled content enhancement:
  - inserts a conservative test-environment/result-analysis subsection when the test chapter is thin;
  - inserts a generic acknowledgements section when missing.
- Added iterative static TOC synchronization:
  - renders the DOCX to PDF;
  - reads real heading page labels;
  - rewrites static TOC entries;
  - repeats until page labels converge or the pass limit is reached.
- Added `.doc` conversion safeguards:
  - removes partial duplicate front matter before inserting the complete template front matter;
  - places the body section break after TOC content controls, not after the TOC title.
- Added a bounded diagram repair step for recognizable simple system architecture flowcharts, replacing broken connector images with a clean generated diagram.

## Verification

Commands:

```bash
PYTHONPATH=src pytest -q
PYTHONPATH=src python3 -m thesis_agent batch-process \
  --template /path/to/private-format-template.doc \
  --inputs /path/to/private-drafts \
  --out runs/acceptance-redcheck
```

Results:

| Draft | Gate | Score | Blockers |
| --- | --- | ---: | --- |
| 私有样本 A | PASS | 96 | 0 |
| 私有样本 B | PASS | 96 | 0 |
| 私有样本 C | PASS | 96 | 0 |
| 私有样本 D | PASS | 100 | 0 |
| 私有样本 E | PASS | 100 | 0 |

The latest generated outputs are under `runs/acceptance-redcheck/`.

## Remaining Work

- The visual package is still generated for external VLM review; automatic ingestion of VLM comments into another edit pass remains future work.
- Reference formatting is counted and warned on, but automatic GB/T 7714 normalization is still not implemented.
