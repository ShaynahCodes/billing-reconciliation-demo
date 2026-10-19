# Changelog

## Portfolio edition

### 2026-10
- Regression suite expanded from 5 to 15 tests, covering every validation rule.
- Tests run automatically on Windows (Python 3.10 and 3.12) through GitHub Actions.
- Documentation reorganized into the README, `docs/USER_GUIDE.md` and this changelog.

### v2.3.3
- The four workflow buttons stay in one row at every supported window size. Button widths are measured against real Tk geometry so Windows display scaling can't push them off screen.

### Initial portfolio release
- Company, customer and email recipients moved out of the code and into `config.json`.
- All real business data replaced with fictional sample data and a demo Ship To library.
- Runtime data (`data/`, `logs/`, generated PDFs) kept out of version control.

## Application history

### v2.3
- Main screen reworked as a checklist. Summary cards hold the totals, and the progress panel says what to do next after each stage.
- Workflow buttons unlock in order: Validate → Generate PDFs → Review PDFs → Outlook Draft.
- Footer and Ship To windows keep their buttons visible when resized.

### v2.2
- Invoice packet and open-balance PDFs rebuilt to an approved visual design, generated directly from the CSV data.
- Consistent grid, number alignment and currency formatting across both PDFs.

### Earlier releases
- Ship To library preloaded from historical billing weeks.
- Safe auto-matching when QuickBooks splits a known address differently, limited to clear unique matches.
- Manage Ship Tos window with search and full multi-line preview.
