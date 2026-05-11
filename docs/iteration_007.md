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
- Enforced the 80-point quality gate: warning-only issues still reduce the score, and any score below 80 fails even without other hard blockers.
- Hardened slot filling around figures:
  - preserves an embedded visual when a source paragraph contains both the image and figure caption text;
  - preserves orphan figure captions when the source document has no corresponding image, so missing figures remain visible as student返修 blockers instead of being silently hidden or generated.
- Added conservative content enhancement for figure explanations, thin testing/summary chapters, repeated words, and punctuation cleanup.
- Added per-paper `revision_checklist.md` output listing blockers, figure issues, content issues, and the processing principle that missing figures must be supplied by the student.

## Acceptance

The full sample set was reproduced under `runs/acceptance-score80-final-v2`.

- `毕业论文.docx`: PASS, score 80, blockers 0
- `物联网2212-杨钰婷-毕业论文初稿.docx`: PASS, score 80, blockers 0
- `物联网2212-蔡宇璐-冷链物流温控追踪系统-初稿.docx`: PASS, score 92, blockers 0
- `论文2.doc`: PASS, score 80, blockers 0
- `论文初.doc`: FAIL, score 81, blocker: source document contains figure captions without corresponding images

All passing audits report no blank pages, no front-matter Arabic page-number errors, no TOC page-number mismatches, and no abstract, TOC, heading, body paragraph, caption, reference, or acknowledgement format errors. The failed paper is intentionally not auto-passed because missing figures must be supplied by the student.
