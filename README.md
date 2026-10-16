# Billing Reconciliation Demo | Python Billing Reconciliation Automation

A desktop workflow automation built with **Python and Tkinter** that reconciles weekly QuickBooks CSV exports, checks invoice and shipment details, and generates two review-ready PDF reports. Designed for a human-in-the-loop billing workflow rather than automatically transmitting invoices.

> **Portfolio demo:** All company names, customers, emails, invoice numbers, PO numbers, shipping sites, and sample billing records are fictional. This edition uses a clean demo shipping library and keeps runtime state out of the repository. It does not log into, read from, or write to QuickBooks.

## Application preview

### 1. Before validation

The application starts with empty QuickBooks export fields and billing actions waiting for the CSV files.

![Billing Reconciliation Demo before validation, with no files selected](docs/app-before-validation.png)

### 2. After validation

With the included fictional CSV exports, the app reconciles **35 invoices** and displays a **$5,265.07 balance due** with **no validation issues**, enabling the next step, **Generate Billing PDFs**.

![Billing Reconciliation Demo after validating 35 invoices](docs/app-after-validation.png)

## What it does

1. **Reads and reconciles two exports:** compares invoice detail with the open-balance report, including missing invoices, mismatched totals, incorrect dates, and excluded paid invoices.
2. **Validates business rules:** checks packing/handling charges, tax lines, and FedEx tracking information. Incorrect totals produce errors; missing tracking produces warnings for review.
3. **Resolves shipping addresses:** reconstructs Ship To information using a user-maintained verified address library. It only attempts an automatic historical match when evidence is strong and unique; ambiguous cases require a person.
4. **Generates billing PDFs:** creates a multi-page invoice packet and a balance summary using a built-in Python PDF writer.
5. **Prepares an Outlook draft:** opens a prefilled Outlook on the web draft for review, with example.com demo recipients. A person must attach files and send the message manually.
6. **Guards against repeat billing:** warns when generating PDFs for a period already recorded in the local billing history.

## Example output

| Invoice packet | Open-balance summary |
|---|---|
| ![First page of fictional invoice packet](docs/sample-invoice.png) | ![Fictional weekly open-balance summary](docs/sample-balance.png) |

These screenshots were rendered from the included fictional CSV files. The sample run generates **35 invoice pages** plus a **1-page balance report**.

## Quick start (Windows)

**Requirements:** Python 3.10+ with Tkinter. The application and tests rely on the Python standard library. Windows is the intended platform for desktop use and Outlook integration.

1. Download or clone this repository and extract it to a writable folder.
2. Double-click **`Billing Reconciliation Demo.pyw`** or start it from PowerShell:
   ```powershell
   py "Billing Reconciliation Demo.pyw"
   ```
3. Browse to the two files in `TEST_DATA/`:
   - `TEST Invoice Data 09-14-2026 to 09-18-2026.csv`
   - `TEST Open Balance 09-14-2026 to 09-18-2026.csv`
4. Click **VALIDATE BILLING**. Expected results: **35 found**, **35 included**, **$5,265.07 balance**, **$0.00 payments/credits**, **zero errors**, **zero warnings**.
5. Click **GENERATE BILLING PDFs** and choose a folder to save the PDFs. Use **REVIEW PDFs** to inspect the results.
6. The Outlook draft button is for demonstration only. It opens a draft addressed to fictional `example.com` recipients. **Do not send real mail from the demo.**

**Why the Invoice Data CSV has ~142 data rows:** each of the 35 invoices has four export rows (header, product, packing, sales tax). There are also a report-period row and a closing row. With the column header, the file has 143 total CSV rows.

## Run automated tests

From the repository folder, run:

```powershell
py -m unittest discover -s tests -v
```

Five tests cover:

- Successful reconciliation of the 35-invoice fictional dataset
- Hard-error detection for an invoice amount mismatch
- Warning detection for an invoice with missing tracking information
- Resolution of the test week's shipping addresses
- Creation and basic structural checks of both generated PDF files

The tests use temporary modified CSVs for failure cases and don't send email. The GUI layout was also checked at several simulated desktop sizes (including 1024px laptops) and all workflow buttons remain visible in one row. Test interaction on your own Windows PC before public release.

## Repository structure

```text
billing-reconciliation-demo/
├── Billing Reconciliation Demo.pyw     # Tkinter interface and automation logic
├── config.json                       # Fictional demo business rules and recipients
├── TEST_DATA/
│   ├── TEST Invoice Data ... .csv     # 35 sample invoices (QuickBooks export format)
│   ├── TEST Open Balance ... .csv     # Matching fictional open balances
│   └── demo_ship_to_master.json       # Clean shipping-address fixture
├── tests/
│   └── test_billing.py                # Regression tests (unittest)
├── docs/
│   ├── app-before-validation.png      # Initial Windows app screenshot
│   ├── app-after-validation.png       # Validated Windows app screenshot
│   ├── sample-invoice.png             # Rendered invoice sample
│   ├── sample-balance.png             # Rendered balance sample
│   └── PUBLISH_CHECKLIST.md           # Final release checklist
├── START_HERE.txt                     # Detailed workflow notes
├── CHANGELOG_v2_3.txt                 # Original version history
└── .gitignore                         # Runtime output and machine-local files
```

## Local data and safety

- On first launch, the program copies `TEST_DATA/demo_ship_to_master.json` into `data/ship_to_master.json`. Future address edits affect **only this local file**; the clean fixture stays unchanged.
- `data/` and `logs/` are ignored by Git. These include mutable shipping addresses, generation history, and diagnostic logs.
- Generated PDFs, exports, environment files, and common temporary folders are also excluded by `.gitignore`.
- The GUI never changes QuickBooks data, and **never sends an email automatically**.
- If you use this project with real business data, keep all sensitive exports and customer details out of a public repository.

## Tech stack

**Python**, **Tkinter / ttk**, **csv**, **decimal**, **dataclasses**, **json**, **regex**, **pathlib**, **unittest**, and a minimal built-in PDF renderer. No third-party pip packages are required to run the application.

## Development context

This is a privacy-safe portfolio adaptation of a working internal billing automation, published with permission to share the source after removing identifying business details. It demonstrates CSV reconciliation, rule-based exceptions, user-assisted review, PDF generation, desktop GUI engineering, and release hygiene.
