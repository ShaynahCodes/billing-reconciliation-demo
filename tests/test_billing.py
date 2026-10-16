"""Standard-library regression tests for the fictional billing portfolio demo.

Run: py -m unittest discover -s tests -v
"""
import csv
import importlib.machinery
import importlib.util
import re
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
import sys
sys.modules[spec.name] = billing
loader.exec_module(billing)

DETAIL = ROOT / 'TEST_DATA' / 'TEST Invoice Data 09-14-2026 to 09-18-2026.csv'
BALANCE = ROOT / 'TEST_DATA' / 'TEST Open Balance 09-14-2026 to 09-18-2026.csv'


class BillingDemoTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.cfg = billing.load_config()

    def run_validation(self, detail=DETAIL, balance=BALANCE):
        # Don't add test-run entries to the application's log file.
        with mock.patch.object(billing, 'technical_log'):
            return billing.validate_files(detail, balance, self.cfg)

    def test_clean_sample_reconciles_35_invoices(self):
        result = self.run_validation()
        self.assertEqual(len(result.invoices), 35)
        self.assertEqual(len(result.include_nums), 35)
        self.assertEqual(result.open_total, Decimal('5265.07'))
        self.assertEqual(result.payments_total, Decimal('0.00'))
        self.assertEqual(result.hard_errors, [])
        self.assertEqual(result.warnings, [])

    def test_mismatched_invoice_amount_is_blocked(self):
        with tempfile.TemporaryDirectory() as tmp:
            altered = Path(tmp) / 'wrong_balance.csv'
            with BALANCE.open(newline='', encoding='utf-8-sig') as f:
                rows = list(csv.reader(f))
            for row in rows:
                if len(row) >= 8 and row[1] == 'Invoice':
                    row[7] = '999.99'  # export's original Amount field
                    break
            with altered.open('w', newline='', encoding='utf-8-sig') as f:
                csv.writer(f).writerows(rows)
            result = self.run_validation(balance=altered)
            self.assertIn('ORIGINAL_AMOUNT_MISMATCH', {x.code for x in result.hard_errors})

    def test_missing_tracking_produces_warning(self):
        with tempfile.TemporaryDirectory() as tmp:
            altered = Path(tmp) / 'no_tracking.csv'
            # The sample is Windows-1252; retain its original encoding.
            with DETAIL.open(newline='', encoding='cp1252') as f:
                rows = list(csv.reader(f))
            header = rows[0]
            invoice_col = header.index('Num')
            memo_col = header.index('Memo')
            affected = False
            for row in rows[1:]:
                if row[invoice_col] == '104210' and re.search(r'\d{4}\s+\d{4}\s+\d{4}',row[memo_col]):
                    row[memo_col] = 'Packing & Handling Charges; tracking pending'
                    affected = True
            self.assertTrue(affected, 'Test fixture did not contain the expected tracking memo')
            with altered.open('w', newline='', encoding='cp1252') as f:
                csv.writer(f).writerows(rows)
            result = self.run_validation(detail=altered)
            self.assertTrue(any(x.invoice == '104210' and x.code == 'TRACKING_NOT_FOUND'
                                for x in result.warnings))
            self.assertEqual(result.hard_errors, [])

    def test_demo_ship_to_library_resolves_every_invoice(self):
        result = self.run_validation()
        store = billing.ShipToStore()
        unmatched = []
        for n in result.include_nums:
            invoice = result.invoices[n]
            matched = store.get(invoice.ship_signature) or store.find_unique_match(invoice.ship_raw)
            if not matched:
                unmatched.append(n)
        self.assertEqual(unmatched, [])

    def test_generates_two_readable_pdf_documents(self):
        result = self.run_validation()
        store = billing.ShipToStore()
        for n in result.include_nums:
            inv = result.invoices[n]
            inv.ship_to_block = store.get(inv.ship_signature) or store.find_unique_match(inv.ship_raw)
            self.assertTrue(inv.ship_to_block)
        with tempfile.TemporaryDirectory() as tmp:
            packet = Path(tmp) / 'demo_invoices.pdf'
            balance = Path(tmp) / 'demo_balance.pdf'
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
