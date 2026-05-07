# Iteration 007 - Red-Text Hard Gates

## Changes

- Converted more template red-text requirements into blocking DOCX checks:
  - abstract and keyword formatting;
  - static TOC title, TOC levels, TOC entry fonts, and TOC page labels;
  - main chapter headings;
  - second- and third-level headings;
  - normal body paragraphs;
  - figure and table captions;
  - reference-list numbering and hanging indent;
  - acknowledgement title and body.
- Hardened TOC synchronization:
  - removes third-level TOC entries;
  - formats first-level TOC entries as black 14 pt text;
  - formats second-level TOC entries as black 12 pt text;
  - preserves dotted leaders and updates rendered page labels.
- Fixed a rule conflict where figure-caption keep rules could re-add `keepNext`/`keepLines` to normal body paragraphs containing anchored drawings.
- Rebuilt the known system architecture diagram with a cleaner non-crossing layout for visual inspection.

## Acceptance

The full sample set was processed under `runs/acceptance-red-hardgate`.

- `毕业论文.docx`: PASS, score 96
- `物联网2212-杨钰婷-毕业论文初稿.docx`: PASS, score 96
- `物联网2212-蔡宇璐-冷链物流温控追踪系统-初稿.docx`: PASS, score 96
- `论文2.doc`: PASS, score 100
- `论文初.doc`: PASS, score 100

All rendered audits report no blank pages, no front-matter Arabic page-number errors, no TOC page-number mismatches, and no abstract, TOC, heading, body paragraph, caption, reference, or acknowledgement format errors.
