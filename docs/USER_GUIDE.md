# User Guide

How to run a weekly billing cycle with Billing Reconciliation Demo. All company, customer, address, contact, pricing, invoice, PO and tracking data in this repository is fictional sample data.

## Open the app

Double-click **`Billing Reconciliation Demo.pyw`**. No command prompt is needed.

## The two QuickBooks exports

| Screen label | What the export contains |
|---|---|
| **Invoice Data** | Everything needed to rebuild each invoice: invoice number, date, PO, S.O. number, items, quantities, rates, tracking text and Ship To fields. |
| **Open Balance** | Which invoices still have money due, and the current unpaid balance for each. |

The summary cards after validation:

| Card | Meaning |
|---|---|
| **Invoices Found** | Invoices in the Invoice Data export. |
| **Invoices Included** | Invoices that are still open and belong in this billing run. |
| **Balance Due** | Sum of open balances for included invoices. |
| **Payments / Credits** | Original totals minus open balances, so partial payments are visible. |

## Workflow

The bottom bar holds one utility action and four workflow steps:

**Manage Ship Tos** · **Validate Billing** → **Generate Billing PDFs** → **Review PDFs** → **Create Outlook Draft**

| Stage | Buttons available |
|---|---|
| Before validation | Manage Ship Tos, Validate Billing |
| After a clean validation | Generate Billing PDFs unlocks |
| After PDFs are created | Review PDFs and Create Outlook Draft unlock |

The Billing Progress panel starts blank and fills in one check at a time. After each stage it says what to do next.

1. **Validate Billing.** Browse to both exports and validate. Hard errors block the run; warnings are shown for review.
2. **Resolve Ship Tos, if asked.** If an address can't be matched with confidence, the app names the invoice to open in QuickBooks and asks for the full Ship To block.
3. **Add missing tracking, if asked.** Paste the FedEx tracking number found in the customer's order email thread.
4. **Generate Billing PDFs.** Choose a folder. The app creates the invoice packet and the open-balance summary. If PDFs were already generated for this billing week, it asks before generating again.
5. **Review PDFs** before anything is sent.
6. **Create Outlook Draft.** Opens a prefilled Outlook on the web draft. A person attaches the PDFs and sends it.

## Ship To library

Ship To addresses are stored as complete multi-line blocks with line breaks preserved. **Manage Ship Tos** lets you search the library and preview the exact block the app will print.

QuickBooks sometimes splits the same address across its columns differently from week to week. The app auto-matches a known address only when one saved block is a clear, unique match. Two sites with the same company name and street are left for a person to choose.

On first launch, the clean demo library at `TEST_DATA/demo_ship_to_master.json` is copied to `data/ship_to_master.json`. Your edits only change that local copy.

## First test run

Use the two files in `TEST_DATA/` for 09/14/2026 – 09/18/2026. Expected result:

- 35 invoices found, 35 included
- $5,265.07 balance due, $0.00 payments/credits
- 0 hard errors, 0 warnings
- All 13 Ship To signatures resolve without a prompt: 11 by direct lookup and 2 through the unique-match path

## Configuration

All business details live in `config.json`: company header and footer, Bill To block, expected packing item and rate, tax label, PDF filename prefixes and Outlook To/CC recipients.

## Local files

- `data/` holds your Ship To library copy and billing history.
- `logs/` holds diagnostic logs.
- Both folders and all generated PDFs are ignored by Git.
