"""Standard-library regression tests for the fictional billing portfolio demo.

Each failure-case test copies a clean sample export, changes one value the way a
real QuickBooks mistake would, and checks that the bot reports the right issue
at the right severity:

    HARD     blocks PDF generation until fixed
    WARNING  allowed through, but flagged for a person to review
    INFO     expected behavior the reviewer should know about

Run: py -m unittest discover -s tests -v
"""
import csv
import importlib.machinery
import importlib.util
import shutil
import sys
import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / 'Billing Reconciliation Demo.pyw'
loader = importlib.machinery.SourceFileLoader('billing_demo', str(SOURCE))
spec = importlib.util.spec_from_loader(loader.name, loader)
billing = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = billing
loader.exec_module(billing)

DETAIL = ROOT / 'TEST_DATA' / 'TEST Invoice Data 09-14-2026 to 09-18-2026.csv'
BALANCE = ROOT / 'TEST_DATA' / 'TEST Open Balance 09-14-2026 to 09-18-2026.csv'

# Column positions in the Open Balance export: ,Type,Date,Num,Memo,Due Date,Open Balance,Amount
BAL_NUM, BAL_OPEN, BAL_AMOUNT = 3, 6, 7

SAMPLE_INVOICE = '104210'
SAMPLE_TOTAL = Decimal('224.25')


def read_rows(path, encoding):
    with path.open(newline='', encoding=encoding) as f:
        return list(csv.reader(f))


def write_rows(path, rows, encoding):
    with path.open('w', newline='', encoding=encoding) as f:
        csv.writer(f).writerows(rows)


class BillingDemoTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cfg = billing.load_config()
        # Keep tests away from the real data/ and logs/ folders.
        cls._sandbox = tempfile.TemporaryDirectory()
        sandbox = Path(cls._sandbox.name)
        cls._patches = [
            mock.patch.object(billing, 'DATA_DIR', sandbox / 'data'),
            mock.patch.object(billing, 'LOG_DIR', sandbox / 'logs'),
            mock.patch.object(billing, 'SHIP_MASTER_PATH', sandbox / 'data' / 'ship_to_master.json'),
            mock.patch.object(billing, 'HISTORY_PATH', sandbox / 'data' / 'billing_history.json'),
        ]
        for p in cls._patches:
            p.start()

    @classmethod
    def tearDownClass(cls):
        for p in cls._patches:
            p.stop()
        cls._sandbox.cleanup()

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        # Fresh copy of the demo address library for every test.
        billing.SHIP_MASTER_PATH.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(billing.DEMO_SHIP_MASTER, billing.SHIP_MASTER_PATH)

    # ---- helpers ---------------------------------------------------------

    def run_validation(self, detail=DETAIL, balance=BALANCE):
        with mock.patch.object(billing, 'technical_log'):
            return billing.validate_files(detail, balance, self.cfg)

    def altered_balance(self, change):
        """Copy the Open Balance CSV, apply change(rows), return the new path."""
        rows = read_rows(BALANCE, 'utf-8-sig')
        change(rows)
        path = Path(self.tmp.name) / 'balance.csv'
        write_rows(path, rows, 'utf-8-sig')
        return path

    def altered_detail(self, change):
        """Copy the Invoice Data CSV (Windows-1252), apply change(header, rows)."""
        rows = read_rows(DETAIL, 'cp1252')
        change(rows[0], rows[1:])
        path = Path(self.tmp.name) / 'detail.csv'
        write_rows(path, rows, 'cp1252')
        return path

    @staticmethod
    def balance_row(rows, invoice):
        for row in rows:
            if len(row) > BAL_AMOUNT and row[1] == 'Invoice' and row[BAL_NUM] == invoice:
                return row
        raise AssertionError(f'Invoice {invoice} not found in Open Balance fixture')

    @staticmethod
    def detail_rows(header, rows, invoice, predicate):
        num = header.index('Num')
        found = [r for r in rows if len(r) > num and r[num] == invoice and predicate(r)]
        if not found:
            raise AssertionError(f'Invoice {invoice} fixture row not found')
        return found

    @staticmethod
    def codes(issues, invoice=None):
        return {x.code for x in issues if invoice is None or x.invoice == invoice}

    # ---- clean run -------------------------------------------------------

    def test_clean_sample_reconciles_35_invoices(self):
        result = self.run_validation()
        self.assertEqual(len(result.invoices), 35)
        self.assertEqual(len(result.include_nums), 35)
        self.assertEqual(result.open_total, Decimal('5265.07'))
        self.assertEqual(result.payments_total, Decimal('0.00'))
        self.assertEqual(result.hard_errors, [])
        self.assertEqual(result.warnings, [])

    # ---- reconciliation between the two exports ---------------------------

    def test_mismatched_invoice_amount_is_blocked(self):
        def change(rows):
            self.balance_row(rows, SAMPLE_INVOICE)[BAL_AMOUNT] = '999.99'
        result = self.run_validation(balance=self.altered_balance(change))
        self.assertIn('ORIGINAL_AMOUNT_MISMATCH', self.codes(result.hard_errors, SAMPLE_INVOICE))

    def test_open_balance_above_invoice_total_is_blocked(self):
        def change(rows):
            self.balance_row(rows, SAMPLE_INVOICE)[BAL_OPEN] = '300.00'
        result = self.run_validation(balance=self.altered_balance(change))
        self.assertIn('OPEN_GREATER_THAN_ORIGINAL', self.codes(result.hard_errors, SAMPLE_INVOICE))

    def test_open_invoice_missing_from_invoice_data_is_blocked(self):
        def change(rows):
            extra = list(self.balance_row(rows, SAMPLE_INVOICE))
            extra[BAL_NUM] = '109999'
            rows.insert(3, extra)
        result = self.run_validation(balance=self.altered_balance(change))
        self.assertIn('OPEN_INVOICE_MISSING_DETAIL', self.codes(result.hard_errors, '109999'))

    def test_paid_invoice_is_excluded_from_the_run(self):
        def change(rows):
            self.balance_row(rows, SAMPLE_INVOICE)[BAL_OPEN] = '0.00'
        result = self.run_validation(balance=self.altered_balance(change))
        self.assertNotIn(SAMPLE_INVOICE, result.include_nums)
        self.assertEqual(len(result.include_nums), 34)
        self.assertIn('DETAIL_NOT_OPEN', self.codes(result.infos, SAMPLE_INVOICE))
        self.assertEqual(result.hard_errors, [])
        self.assertEqual(result.open_total, Decimal('5265.07') - SAMPLE_TOTAL)

    def test_partial_payment_is_calculated(self):
        def change(rows):
            self.balance_row(rows, SAMPLE_INVOICE)[BAL_OPEN] = '100.00'
        result = self.run_validation(balance=self.altered_balance(change))
        self.assertEqual(result.hard_errors, [])
        self.assertIn(SAMPLE_INVOICE, result.include_nums)
        self.assertEqual(result.invoices[SAMPLE_INVOICE].payments_credits, SAMPLE_TOTAL - Decimal('100.00'))
        self.assertEqual(result.payments_total, Decimal('124.25'))
        self.assertEqual(result.open_total, Decimal('5265.07') - Decimal('124.25'))

    def test_credit_memo_requires_manual_review(self):
        def change(rows):
            memo = list(self.balance_row(rows, SAMPLE_INVOICE))
            memo[1], memo[BAL_NUM] = 'Credit Memo', '50001'
            rows.insert(3, memo)
        result = self.run_validation(balance=self.altered_balance(change))
        self.assertIn('CREDIT_MEMO_FOUND', self.codes(result.hard_errors))

    # ---- business rules on each invoice -----------------------------------

    def test_wrong_packing_rate_is_blocked(self):
        def change(header, rows):
            col = header.index('Sales Price')
            for r in self.detail_rows(header, rows, SAMPLE_INVOICE,
                                      lambda r: 'Packing and Handling' in r[header.index('Item')]):
                r[col] = '15.00'
        result = self.run_validation(detail=self.altered_detail(change))
        self.assertIn('WRONG_PACKING_RATE', self.codes(result.hard_errors, SAMPLE_INVOICE))

    def test_unexpected_sales_tax_produces_warning(self):
        def change(header, rows):
            col = header.index('Amount')
            for r in self.detail_rows(header, rows, SAMPLE_INVOICE,
                                      lambda r: r[header.index('Account')] == 'Sales Tax Payable'):
                r[col] = '13.46'
        result = self.run_validation(detail=self.altered_detail(change))
        self.assertIn('NONZERO_SALES_TAX', self.codes(result.warnings, SAMPLE_INVOICE))
        self.assertEqual(result.hard_errors, [])

    def test_missing_tracking_produces_warning(self):
        def change(header, rows):
            memo = header.index('Memo')
            for r in self.detail_rows(header, rows, SAMPLE_INVOICE, lambda r: 'Tracking' in r[memo]):
                r[memo] = 'Packing & Handling Charges; tracking pending'
        result = self.run_validation(detail=self.altered_detail(change))
        self.assertIn('TRACKING_NOT_FOUND', self.codes(result.warnings, SAMPLE_INVOICE))
        self.assertEqual(result.hard_errors, [])

    def test_invoice_dated_outside_billing_week_is_blocked(self):
        def change(header, rows):
            col = header.index('Date')
            for r in self.detail_rows(header, rows, SAMPLE_INVOICE, lambda r: True):
                r[col] = '09/21/2026'
        result = self.run_validation(detail=self.altered_detail(change))
        self.assertIn('DATE_OUTSIDE_PERIOD', self.codes(result.hard_errors, SAMPLE_INVOICE))

    # ---- duplicate billing guard -----------------------------------------

    def test_previously_billed_week_is_detected(self):
        result = self.run_validation()
        label = billing.period_filename(result.period)
        self.assertEqual(billing.prior_runs_for_period(billing.load_history(), label), [])

        billing.save_history([
            {'period': '09-07-2026 to 09-11-2026', 'generated_at': '2026-09-14T09:00:00'},
            {'period': label, 'generated_at': '2026-09-21T09:15:00'},
        ])
        prior = billing.prior_runs_for_period(billing.load_history(), label)
        self.assertEqual([p['generated_at'] for p in prior], ['2026-09-21T09:15:00'])

    # ---- Ship To resolution ----------------------------------------------

    def test_demo_ship_to_library_resolves_every_invoice(self):
        result = self.run_validation()
        store = billing.ShipToStore()
        unmatched = [n for n in result.include_nums
                     if not (store.get(result.invoices[n].ship_signature)
                             or store.find_unique_match(result.invoices[n].ship_raw))]
        self.assertEqual(unmatched, [])

    def test_ambiguous_ship_to_is_left_for_a_person(self):
        # Two saved Harborline sites share a company name and street address.
        # Without the phone/area-code clue the bot must not pick one.
        store = billing.ShipToStore()
        ambiguous = {
            'Ship To Address 1__2': 'Harborline Cold Storage',
            'Ship To Address 2': '4100 Bayou Parkway',
        }
        self.assertEqual(store.find_unique_match(ambiguous), '')

        distinct = dict(ambiguous, **{'Ship To City': 'Mobile: +1 (985)', 'Ship Zip': '555-0174'})
        self.assertIn('ATTE:Sam Patel', store.find_unique_match(distinct))

    # ---- output ----------------------------------------------------------

    def test_generates_two_readable_pdf_documents(self):
        result = self.run_validation()
        store = billing.ShipToStore()
        for n in result.include_nums:
            inv = result.invoices[n]
            inv.ship_to_block = store.get(inv.ship_signature) or store.find_unique_match(inv.ship_raw)
            self.assertTrue(inv.ship_to_block)
        packet = Path(self.tmp.name) / 'demo_invoices.pdf'
        balance = Path(self.tmp.name) / 'demo_balance.pdf'
        billing.generate_invoice_packet(result, self.cfg, packet)
        billing.generate_open_balance(result, self.cfg, balance)
        for document in (packet, balance):
            data = document.read_bytes()
            self.assertTrue(data.startswith(b'%PDF-'), f'{document.name} missing PDF header')
            self.assertIn(b'%%EOF', data[-100:])
            self.assertGreater(len(data), 1000)
        self.assertGreater(packet.stat().st_size, balance.stat().st_size)


if __name__ == '__main__':
    unittest.main()
