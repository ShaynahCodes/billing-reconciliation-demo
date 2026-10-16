# -*- coding: utf-8 -*-
"""
Billing Reconciliation Demo v2
Weekly QuickBooks billing-packet automation (portfolio edition)

Features:
- Two-CSV QuickBooks validation
- Friendly dashboard UI
- In-app unknown Ship To review
- In-app Ship To manager
- Pure-Python PDF generation (no third-party packages)
- Duplicate-period generation warning
- Technical logging kept out of the normal user interface

This application NEVER connects to QuickBooks and cannot edit QuickBooks.
"""

from __future__ import annotations

import csv
import json
import os
import re
import sys
import traceback
import time
import urllib.parse
import subprocess
import ctypes
from ctypes import wintypes
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

APP_NAME = "Billing Reconciliation Demo v2.3.1"
BUILD_REVISION = 28  # internal build reference; intentionally not shown in the UI
ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
LOG_DIR = ROOT / "logs"
CONFIG_PATH = ROOT / "config.json"
DEMO_SHIP_MASTER = ROOT / "TEST_DATA" / "demo_ship_to_master.json"
SHIP_MASTER_PATH = DATA_DIR / "ship_to_master.json"
HISTORY_PATH = DATA_DIR / "billing_history.json"

COLORS = {
    "navy": "#17324D",
    "blue": "#24557A",
    "teal": "#2A7F7F",
    "green": "#2E7D55",
    "green_bg": "#EAF6EF",
    "red": "#A73939",
    "red_bg": "#FBECEC",
    "amber": "#A86B17",
    "amber_bg": "#FFF4DD",
    "light": "#F5F7FA",
    "card": "#FFFFFF",
    "border": "#D9E0E7",
    "text": "#1F2933",
    "muted": "#65727E",
}

ZERO = Decimal("0.00")

DEFAULT_INVOICE_MESSAGE = "Thank you for your business. We appreciate it very much."


@dataclass
class Issue:
    level: str
    code: str
    message: str
    invoice: str = ""


@dataclass
class Invoice:
    num: str
    date: str
    po: str
    customer: str
    customer_address: str
    so_num: str
    via: str
    terms: str
    due_date: str
    original_total: Decimal
    open_balance: Decimal = ZERO
    balance_report_amount: Decimal = ZERO
    payments_credits: Decimal = ZERO
    rows: List[Dict[str, str]] = field(default_factory=list)
    ship_signature: str = ""
    ship_raw: Dict[str, str] = field(default_factory=dict)
    ship_to_block: str = ""
    tracking_override: str = ""
    tracking_detected_text: str = ""
    customer_message: str = ""


@dataclass
class ValidationResult:
    detail_path: Path
    balance_path: Path
    invoices: Dict[str, Invoice]
    balances: Dict[str, Dict[str, object]]
    include_nums: List[str]
    issues: List[Issue]
    period: Optional[Tuple[date, date]]
    original_total: Decimal
    open_total: Decimal
    payments_total: Decimal
    detail_encoding: str
    balance_encoding: str

    @property
    def hard_errors(self):
        return [x for x in self.issues if x.level == "HARD"]

    @property
    def warnings(self):
        return [x for x in self.issues if x.level == "WARNING"]

    @property
    def infos(self):
        return [x for x in self.issues if x.level == "INFO"]


def ensure_files():
    DATA_DIR.mkdir(exist_ok=True)
    LOG_DIR.mkdir(exist_ok=True)
    if not SHIP_MASTER_PATH.exists():
        # Keep the committed demo address library separate from mutable user data.
        # First launch gets an editable local copy; subsequent edits are preserved.
        if DEMO_SHIP_MASTER.exists():
            SHIP_MASTER_PATH.write_bytes(DEMO_SHIP_MASTER.read_bytes())
        else:
            SHIP_MASTER_PATH.write_text("{}", encoding="utf-8")
    if not HISTORY_PATH.exists():
        HISTORY_PATH.write_text("[]", encoding="utf-8")


def load_config():
    return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))


def technical_log(message: str):
    ensure_files()
    path = LOG_DIR / f"Billing_{datetime.now():%Y-%m-%d}.log"
    with path.open("a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%H:%M:%S}] {message}\n")


def detect_encoding(path: Path) -> str:
    raw = path.read_bytes()[:100000]
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            raw.decode(enc)
            return enc
        except UnicodeDecodeError:
            pass
    return "cp1252"


def dedupe_headers(headers: List[str]) -> List[str]:
    seen = defaultdict(int)
    out = []
    for h in headers:
        h = (h or "").strip() or "__col0"
        seen[h] += 1
        out.append(h if seen[h] == 1 else f"{h}__{seen[h]}")
    return out


def read_qb_csv(path: Path):
    enc = detect_encoding(path)
    with path.open("r", newline="", encoding=enc, errors="replace") as f:
        reader = csv.reader(f)
        raw_headers = next(reader, None)
        if not raw_headers:
            raise ValueError(f"{path.name} is empty.")
        headers = dedupe_headers(raw_headers)
        rows = []
        for vals in reader:
            vals = vals[:len(headers)] + [""] * max(0, len(headers) - len(vals))
            rows.append({h: (v or "").strip() for h, v in zip(headers, vals)})
    return rows, headers, enc


def get(row, *names):
    for n in names:
        if n in row and str(row[n]).strip():
            return str(row[n]).strip()
    return ""


def dec(value) -> Optional[Decimal]:
    s = str(value or "").strip()
    if not s or s.endswith("%"):
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    s = s.replace("$", "").replace(",", "").strip()
    try:
        x = Decimal(s)
        return -x if neg else x
    except InvalidOperation:
        return None


def money(x: Decimal) -> str:
    return f"${x.quantize(Decimal('0.01')):,.2f}"


def invoice_num(value) -> str:
    s = str(value or "").strip()
    return s[:-2] if re.fullmatch(r"\d+\.0", s) else s


def parse_date(s: str) -> Optional[date]:
    for fmt in ("%m/%d/%Y", "%m/%d/%y"):
        try:
            return datetime.strptime((s or "").strip(), fmt).date()
        except ValueError:
            pass
    return None


MONTHS = {m.lower(): i for i, m in enumerate(
    ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"], start=1
)}


def parse_report_period(rows) -> Optional[Tuple[date, date]]:
    candidates = [get(r, "__col0") for r in rows[:6]]
    for s in candidates:
        m = re.fullmatch(r"([A-Za-z]{3})\s+(\d{1,2})\s*-\s*(\d{1,2}),\s*(\d{2})", s)
        if m:
            mon, d1, d2, yy = m.groups()
            y = 2000 + int(yy)
            mo = MONTHS.get(mon.lower())
            if mo:
                return date(y, mo, int(d1)), date(y, mo, int(d2))
        m = re.fullmatch(r"([A-Za-z]{3})\s+(\d{1,2})\s*-\s*([A-Za-z]{3})\s+(\d{1,2}),\s*(\d{2})", s)
        if m:
            mon1, d1, mon2, d2, yy = m.groups()
            y = 2000 + int(yy)
            mo1, mo2 = MONTHS.get(mon1.lower()), MONTHS.get(mon2.lower())
            if mo1 and mo2:
                return date(y, mo1, int(d1)), date(y, mo2, int(d2))
    return None


def is_header_row(r):
    return get(r, "Split") == "-SPLIT-" or (
        get(r, "Account").lower() == "accounts receivable"
        and not get(r, "Item")
        and dec(get(r, "Open Balance")) is not None
    )


def original_total(header):
    for col in ("Amount", "Debit", "Open Balance"):
        x = dec(get(header, col))
        if x is not None:
            return abs(x)
    return ZERO


def ship_fields(row):
    keys = [k for k in row if k.lower().startswith("ship to") or k.lower() == "ship zip"]
    return {k: get(row, k) for k in keys if get(row, k)}


def normalize_signature(parts):
    vals = []
    for k in sorted(parts):
        v = re.sub(r"\s+", " ", parts[k].strip()).lower()
        if v:
            vals.append(f"{k.lower()}={v}")
    return " | ".join(vals)


def extract_tracking_number(text):
    """Return a normalized 12-digit shipment number when one is present.

    Handles formats seen in real-world invoice memos such as:
      Tracking # 8771 8504 8205
      Fedex # 8778 6322 0018
      FEDEX#877863220018
      8778-6322-0018
    """
    t = str(text or "")

    # Prefer a 12-digit sequence with optional spaces/hyphens between digits.
    # This avoids treating unrelated dates/amounts elsewhere in the row as
    # tracking while accepting coworkers' normal FedEx formatting.
    for match in re.finditer(r"(?<!\d)(\d(?:[\s-]*\d){11})(?!\d)", t):
        digits = re.sub(r"\D", "", match.group(1))
        if len(digits) == 12:
            return digits
    return ""


def looks_like_tracking(text):
    return bool(extract_tracking_number(text))


def tracking_text_from_row(row):
    """Collect only description-like fields where tracking is expected."""
    fields = [
        get(row, "Memo"),
        get(row, "Item Description"),
        get(row, "Description"),
        get(row, "Line Description"),
    ]
    return "\n".join(v for v in fields if v)


def concise_tracking_text(text):
    """Keep the useful tracking line(s) without copying a giant unrelated field."""
    raw = str(text or "").strip()
    if not raw:
        return ""

    lines = [line.strip() for line in raw.splitlines() if line.strip()]
    if not lines:
        return ""

    for i, line in enumerate(lines):
        if looks_like_tracking(line):
            chosen = [line]
            # Staff often paste the FedEx delivery estimate on the next line.
            if i + 1 < len(lines):
                nxt = lines[i + 1]
                if len(nxt) <= 80:
                    chosen.append(nxt)
            return "\n".join(chosen)

    # If the tracking number is split oddly across newlines, keep the source
    # field when it is reasonably short.
    if looks_like_tracking(raw):
        return raw if len(raw) <= 180 else f"Tracking # {extract_tracking_number(raw)}"
    return ""


def find_invoice_tracking_text(inv):
    """Find tracking anywhere in the invoice export, not only the packing row.

    QuickBooks/coworker entry patterns can place the FedEx text in Memo,
    Item Description, another description field, or another exported column.
    """
    keyword_hits = []
    other_hits = []

    for row in inv.rows:
        for key, value in row.items():
            value = str(value or "").strip()
            if not value or not looks_like_tracking(value):
                continue

            concise = concise_tracking_text(value)
            if not concise:
                continue

            low = value.lower()
            item = (key, concise)
            if "fedex" in low or "fed ex" in low or "tracking" in low:
                keyword_hits.append(item)
            else:
                other_hits.append(item)

    hits = keyword_hits or other_hits
    return hits[0][1] if hits else ""


def parse_detail(path: Path, cfg):
    rows, _, enc = read_qb_csv(path)
    issues = []
    period = parse_report_period(rows)

    by_num = defaultdict(list)
    for r in rows:
        if get(r, "Type").lower() == "invoice":
            n = invoice_num(get(r, "Num"))
            if n:
                by_num[n].append(r)

    invoices = {}
    for n, rs in by_num.items():
        hs = [r for r in rs if is_header_row(r)]
        if not hs:
            issues.append(Issue("HARD", "DETAIL_HEADER_MISSING",
                                "No invoice header row was found in the Invoice Data CSV.", n))
            continue
        if len(hs) > 1:
            issues.append(Issue("HARD", "DETAIL_HEADER_DUPLICATE",
                                f"{len(hs)} invoice header rows were found.", n))
        h = hs[0]
        inv = Invoice(
            num=n,
            date=get(h, "Date"),
            po=get(h, "P. O. #"),
            customer=get(h, "Name"),
            customer_address=get(h, "Name Address"),
            so_num=invoice_num(get(h, "S. O. #")),
            via=get(h, "Via"),
            terms=get(h, "Terms"),
            due_date=get(h, "Due Date"),
            original_total=original_total(h),
            rows=rs,
            ship_raw=ship_fields(h),
            customer_message=get(
                h,
                "Customer Message",
                "Customer Msg",
                "Message"
            ),
        )
        inv.ship_signature = normalize_signature(inv.ship_raw)
        invoices[n] = inv

        if inv.customer.lower() != cfg["customer_name"].lower():
            issues.append(Issue("HARD", "WRONG_CUSTOMER",
                                f"Expected {cfg['customer_name']}; found {inv.customer or '[blank]'}.", n))

    if period:
        if not (period[0].weekday() == 0 and period[1].weekday() == 4 and (period[1] - period[0]).days == 4):
            issues.append(Issue(
                "HARD", "INVALID_BILLING_PERIOD",
                f"The Invoice Data report is {period[0]:%m/%d/%Y} - {period[1]:%m/%d/%Y}. "
                "Weekly billing should use Monday through Friday."
            ))
        for n, inv in invoices.items():
            d = parse_date(inv.date)
            if d and not (period[0] <= d <= period[1]):
                issues.append(Issue("HARD", "DATE_OUTSIDE_PERIOD",
                                    f"Invoice date {inv.date} is outside the selected billing period.", n))
    else:
        issues.append(Issue("HARD", "PERIOD_NOT_FOUND",
                            "The billing period could not be read from the Invoice Data CSV."))

    return invoices, issues, period, enc


def parse_balances(path: Path):
    rows, _, enc = read_qb_csv(path)
    issues = []
    balances = {}
    for r in rows:
        typ = get(r, "Type").strip()
        if typ.lower() in {"credit memo", "creditmemo", "credit"}:
            issues.append(Issue(
                "HARD", "CREDIT_MEMO_FOUND",
                "A credit memo was found in the Open Balance report. This requires manual review.",
                invoice_num(get(r, "Num"))
            ))
        if typ.lower() != "invoice":
            continue
        n = invoice_num(get(r, "Num"))
        if not n:
            continue
        if n in balances:
            issues.append(Issue("HARD", "DUPLICATE_BALANCE_ROW",
                                "Invoice appears more than once in the Open Balance CSV.", n))
            continue
        balances[n] = {
            "date": get(r, "Date"),
            "due_date": get(r, "Due Date"),
            "open_balance": abs(dec(get(r, "Open Balance")) or ZERO),
            "amount": abs(dec(get(r, "Amount")) or ZERO),
        }
    return balances, issues, enc


def validate_invoice(inv: Invoice, cfg):
    issues = []
    tol = Decimal(str(cfg.get("amount_tolerance", 0.01)))
    expected_item = cfg["expected_packing_item_text"].lower()
    expected_rate = Decimal(str(cfg["expected_packing_rate"]))
    expected_amount = Decimal(str(cfg["expected_packing_amount"]))

    detail_rows = [r for r in inv.rows if not is_header_row(r)]
    packing_rows = []
    tax_rows = []
    for r in detail_rows:
        blob = " ".join([get(r, "Item"), get(r, "Item Description"), get(r, "Memo")]).lower()
        if "packing" in blob and "handling" in blob:
            packing_rows.append(r)
        if "sales tax" in blob or get(r, "Account").lower() == "sales tax payable":
            tax_rows.append(r)

    # Search the entire invoice once. Tracking can be stored on another line
    # or exported into an unexpected QuickBooks column.
    inv.tracking_detected_text = find_invoice_tracking_text(inv)

    if not packing_rows:
        issues.append(Issue("WARNING", "PACKING_LINE_MISSING",
                            "No packing/handling line was found.", inv.num))
    else:
        if len(packing_rows) > 1:
            issues.append(Issue("WARNING", "MULTIPLE_PACKING_LINES",
                                f"{len(packing_rows)} packing/handling lines were found.", inv.num))
        for r in packing_rows:
            item = get(r, "Item")
            rate = dec(get(r, "Sales Price"))
            amount = dec(get(r, "Amount"))
            memo = get(r, "Memo")
            if expected_item not in item.lower():
                issues.append(Issue(
                    "HARD", "WRONG_PACKING_ITEM",
                    f"Packing item is '{item or '[blank]'}'. Billing rules require '{cfg['expected_packing_item_text']}'.",
                    inv.num
                ))
            if rate is None or abs(abs(rate) - expected_rate) > tol:
                issues.append(Issue(
                    "HARD", "WRONG_PACKING_RATE",
                    f"Packing rate is {get(r, 'Sales Price') or '[blank]'}. Billing rules require {money(expected_rate)}.",
                    inv.num
                ))
            if amount is None or abs(abs(amount) - expected_amount) > tol:
                issues.append(Issue(
                    "HARD", "WRONG_PACKING_AMOUNT",
                    f"Packing line amount is {get(r, 'Amount') or '[blank]'}. Billing rules require {money(expected_amount)}.",
                    inv.num
                ))
        if not inv.tracking_detected_text:
            issues.append(Issue(
                "WARNING", "TRACKING_NOT_FOUND",
                "Missing tracking number.",
                inv.num
            ))

    for r in tax_rows:
        amt = dec(get(r, "Amount"))
        if amt is not None and abs(amt) > tol:
            issues.append(Issue("WARNING", "NONZERO_SALES_TAX",
                                f"Sales tax is not zero ({money(abs(amt))}).", inv.num))
    return issues


def validate_files(detail_path: Path, balance_path: Path, cfg) -> ValidationResult:
    invoices, i1, period, detail_enc = parse_detail(detail_path, cfg)
    balances, i2, balance_enc = parse_balances(balance_path)
    issues = i1 + i2
    tol = Decimal(str(cfg.get("amount_tolerance", 0.01)))

    detail_nums = set(invoices)
    open_nums = {n for n, b in balances.items() if b["open_balance"] > ZERO}

    for n in sorted(open_nums - detail_nums):
        issues.append(Issue("HARD", "OPEN_INVOICE_MISSING_DETAIL",
                            "This open invoice is missing from the Invoice Data CSV.", n))

    for n in sorted(detail_nums - open_nums):
        issues.append(Issue("INFO", "DETAIL_NOT_OPEN",
                            "This invoice is in the Detail report but is no longer open, so it will be excluded.", n))

    include = sorted(detail_nums & open_nums, key=lambda x: int(x) if x.isdigit() else x)

    for n in include:
        inv = invoices[n]
        b = balances[n]
        inv.open_balance = b["open_balance"]
        inv.balance_report_amount = b["amount"]
        inv.payments_credits = inv.original_total - inv.open_balance

        if inv.open_balance - inv.original_total > tol:
            issues.append(Issue("HARD", "OPEN_GREATER_THAN_ORIGINAL",
                                f"Open balance {money(inv.open_balance)} is greater than invoice total {money(inv.original_total)}.", n))
        if abs(inv.balance_report_amount - inv.original_total) > tol:
            issues.append(Issue("HARD", "ORIGINAL_AMOUNT_MISMATCH",
                                f"Invoice Data total {money(inv.original_total)} does not match Open Balance Amount {money(inv.balance_report_amount)}.", n))
        dd, bd = parse_date(inv.date), parse_date(str(b["date"]))
        if dd and bd and dd != bd:
            issues.append(Issue("HARD", "INVOICE_DATE_MISMATCH",
                                f"Invoice Data date {inv.date} differs from Open Balance date {b['date']}.", n))
        if period and bd and not (period[0] <= bd <= period[1]):
            issues.append(Issue("HARD", "BALANCE_DATE_OUTSIDE_PERIOD",
                                f"Open Balance invoice date {b['date']} is outside the billing period.", n))
        issues.extend(validate_invoice(inv, cfg))

    original_sum = sum((invoices[n].original_total for n in include), ZERO)
    open_sum = sum((invoices[n].open_balance for n in include), ZERO)
    payments_sum = sum((invoices[n].payments_credits for n in include), ZERO)

    technical_log(
        f"Validated detail={detail_path.name} ({detail_enc}), balance={balance_path.name} ({balance_enc}), "
        f"detail_count={len(invoices)}, include={len(include)}, open_total={open_sum}, "
        f"hard={len([x for x in issues if x.level=='HARD'])}, warnings={len([x for x in issues if x.level=='WARNING'])}"
    )

    return ValidationResult(
        detail_path, balance_path, invoices, balances, include, issues, period,
        original_sum, open_sum, payments_sum, detail_enc, balance_enc
    )


class ShipToStore:
    def __init__(self):
        ensure_files()
        self.data = {}
        self.load()

    def load(self):
        try:
            obj = json.loads(SHIP_MASTER_PATH.read_text(encoding="utf-8"))
            self.data = obj if isinstance(obj, dict) else {}
        except Exception:
            self.data = {}

    def save(self):
        SHIP_MASTER_PATH.write_text(
            json.dumps(self.data, indent=2, ensure_ascii=False),
            encoding="utf-8"
        )

    def get(self, signature):
        rec = self.data.get(signature)
        if isinstance(rec, dict):
            return (rec.get("address_block") or "").strip()
        if isinstance(rec, str):
            return rec.strip()
        return ""

    @staticmethod
    def _norm(value):
        return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())

    def all_unique_blocks(self):
        seen = {}
        for key, rec in self.data.items():
            if isinstance(rec, str):
                block = rec.strip()
                rec = {"address_block": block, "label": block.splitlines()[0] if block else "Saved Ship To"}
            else:
                block = (rec.get("address_block") or "").strip()
            if not block:
                continue
            nk = self._norm(block)
            if nk not in seen:
                seen[nk] = (key, rec)
            elif not rec.get("reference_only", False):
                # Prefer a mapped record over a reference-only library entry.
                seen[nk] = (key, rec)
        return list(seen.values())

    def find_unique_match(self, ship_raw):
        """Try to match a new QuickBooks signature to one historical full block.

        This only auto-matches when the evidence is strong and the best match is
        unique. Ambiguous matches are intentionally left for human review.
        """
        clues = []
        for key, raw in ship_raw.items():
            raw = str(raw or "").strip()
            n = self._norm(raw)
            if len(n) < 4:
                continue
            # State abbreviations / generic labels are too weak on their own.
            if n in {"ga", "tx", "al", "tel", "phone", "mobile"}:
                continue
            weight = 2
            if any(ch.isdigit() for ch in n):
                weight = 3
            if "address 2" in key.lower() or len(n) >= 12:
                weight = 4
            clues.append((n, weight))

        if not clues:
            return ""

        candidates = []
        for _, rec in self.all_unique_blocks():
            block = (rec.get("address_block") or "").strip()
            bn = self._norm(block)
            score = 0
            matched = 0
            for clue, weight in clues:
                if clue and clue in bn:
                    score += weight
                    matched += 1
            if matched:
                candidates.append((score, matched, block))

        if not candidates:
            return ""

        candidates.sort(reverse=True, key=lambda x: (x[0], x[1]))
        best = candidates[0]
        second_score = candidates[1][0] if len(candidates) > 1 else -1

        # Require multiple clues or one very strong compound clue, and clear
        # separation from the runner-up to avoid confusing same-company sites.
        if best[0] >= 6 and best[1] >= 2 and best[0] >= second_score + 2:
            return best[2]
        return ""

    def set(self, signature, block, source="Entered in Billing Bot"):
        block = block.strip()
        self.data[signature] = {
            "address_block": block,
            "label": block.splitlines()[0] if block else "Saved Ship To",
            "source": source,
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "reference_only": False,
        }
        # Also retain the block as a historical/reference address so the bot
        # can match future report signatures that QuickBooks splits differently.
        lib_key = "library::" + __import__("hashlib").sha1(block.encode("utf-8")).hexdigest()[:16]
        if lib_key not in self.data:
            self.data[lib_key] = {
                "address_block": block,
                "label": block.splitlines()[0] if block else "Saved Ship To",
                "source": source,
                "updated_at": datetime.now().isoformat(timespec="seconds"),
                "reference_only": True,
            }
        self.save()

    def delete(self, signature):
        self.data.pop(signature, None)
        self.save()


# -----------------------
# Pure-Python PDF writer
# v2.2 visual spec matches the approved polished invoice/open-balance PDFs.
# -----------------------

def pdf_safe(s: str) -> str:
    b = str(s or "").encode("cp1252", errors="replace")
    s2 = b.decode("latin-1")
    return s2.replace("\\", "\\\\").replace("(", "\\(").replace(")", "\\)")


class SimplePDF:
    def __init__(self):
        self.pages = []

    def add_page(self, commands: List[str]):
        self.pages.append("\n".join(commands))

    def save(self, path: Path):
        objs = []

        def add_obj(content: bytes):
            objs.append(content)
            return len(objs)

        font_regular = add_obj(
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
        )
        font_bold = add_obj(
            b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"
        )

        page_ids = []
        content_ids = []
        for content in self.pages:
            data = content.encode("latin-1", errors="replace")
            content_ids.append(add_obj(
                f"<< /Length {len(data)} >>\nstream\n".encode("ascii")
                + data + b"\nendstream"
            ))

        pages_id_placeholder = len(objs) + len(self.pages) + 1
        for cid in content_ids:
            page_ids.append(add_obj(
                f"<< /Type /Page /Parent {pages_id_placeholder} 0 R "
                f"/MediaBox [0 0 612 792] "
                f"/Resources << /Font << /F1 {font_regular} 0 R /F2 {font_bold} 0 R >> >> "
                f"/Contents {cid} 0 R >>".encode("ascii")
            ))

        kids = " ".join(f"{pid} 0 R" for pid in page_ids)
        pages_id = add_obj(
            f"<< /Type /Pages /Kids [{kids}] /Count {len(page_ids)} >>".encode("ascii")
        )
        catalog_id = add_obj(
            f"<< /Type /Catalog /Pages {pages_id} 0 R >>".encode("ascii")
        )

        out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
        offsets = [0]
        for i, obj in enumerate(objs, start=1):
            offsets.append(len(out))
            out.extend(f"{i} 0 obj\n".encode("ascii"))
            out.extend(obj)
            out.extend(b"\nendobj\n")

        xref = len(out)
        out.extend(f"xref\n0 {len(objs)+1}\n".encode("ascii"))
        out.extend(b"0000000000 65535 f \n")
        for off in offsets[1:]:
            out.extend(f"{off:010d} 00000 n \n".encode("ascii"))
        out.extend(
            f"trailer\n<< /Size {len(objs)+1} /Root {catalog_id} 0 R >>\n"
            f"startxref\n{xref}\n%%EOF\n".encode("ascii")
        )
        path.write_bytes(out)


# Approved palette from the polished FINAL PDFs.
NAVY = (39/255, 76/255, 119/255)       # #274C77
TEXT = (31/255, 41/255, 55/255)        # #1F2937
MUTED = (107/255, 114/255, 128/255)    # #6B7280
BORDER = (209/255, 213/255, 219/255)   # #D1D5DB
ROW_LINE = (229/255, 231/255, 235/255) # #E5E7EB
LIGHT_GRAY = (243/255, 244/255, 246/255) # #F3F4F6
LIGHT_BLUE = (234/255, 241/255, 248/255) # #EAF1F8
WHITE = (1, 1, 1)

# Standard Helvetica ASCII widths, units per 1000 em.
_HELV = {
' ':278,'!':278,'"':355,'#':556,'$':556,'%':889,'&':667,"'":191,'(':333,')':333,'*':389,'+':584,
',':278,'-':333,'.':278,'/':278,'0':556,'1':556,'2':556,'3':556,'4':556,'5':556,'6':556,'7':556,
'8':556,'9':556,':':278,';':278,'<':584,'=':584,'>':584,'?':556,'@':1015,'A':667,'B':667,'C':722,
'D':722,'E':667,'F':611,'G':778,'H':722,'I':278,'J':500,'K':667,'L':556,'M':833,'N':722,'O':778,
'P':667,'Q':778,'R':722,'S':667,'T':611,'U':722,'V':667,'W':944,'X':667,'Y':667,'Z':611,'[':278,
'\\':278,']':278,'^':469,'_':556,'`':333,'a':556,'b':556,'c':500,'d':556,'e':556,'f':278,'g':556,
'h':556,'i':222,'j':222,'k':500,'l':222,'m':833,'n':556,'o':556,'p':556,'q':556,'r':333,'s':500,
't':278,'u':556,'v':500,'w':722,'x':500,'y':500,'z':500,'{':334,'|':260,'}':334,'~':584
}
_HELVB = {
' ':278,'!':333,'"':474,'#':556,'$':556,'%':889,'&':722,"'":238,'(':333,')':333,'*':389,'+':584,
',':278,'-':333,'.':278,'/':278,'0':556,'1':556,'2':556,'3':556,'4':556,'5':556,'6':556,'7':556,
'8':556,'9':556,':':333,';':333,'<':584,'=':584,'>':584,'?':611,'@':975,'A':722,'B':722,'C':722,
'D':722,'E':667,'F':611,'G':778,'H':722,'I':278,'J':556,'K':722,'L':611,'M':833,'N':722,'O':778,
'P':667,'Q':778,'R':722,'S':667,'T':611,'U':722,'V':667,'W':944,'X':667,'Y':667,'Z':611,'[':333,
'\\':278,']':333,'^':584,'_':556,'`':333,'a':556,'b':611,'c':556,'d':611,'e':556,'f':333,'g':611,
'h':611,'i':278,'j':278,'k':556,'l':278,'m':889,'n':611,'o':611,'p':611,'q':611,'r':389,'s':556,
't':333,'u':611,'v':556,'w':778,'x':556,'y':556,'z':500,'{':389,'|':280,'}':389,'~':584
}


def text_width(text, size, bold=False):
    widths = _HELVB if bold else _HELV
    return sum(widths.get(ch, 556) for ch in str(text or "")) * size / 1000.0


def rgb_cmd(rgb, stroke=False):
    op = "RG" if stroke else "rg"
    return f"{rgb[0]:.5f} {rgb[1]:.5f} {rgb[2]:.5f} {op}"


def cmd_text(x, y, text, size=10, bold=False, color=TEXT):
    font = "F2" if bold else "F1"
    return (
        f"{rgb_cmd(color)} BT /{font} {size:.3f} Tf "
        f"{x:.2f} {y:.2f} Td ({pdf_safe(text)}) Tj ET"
    )


def cmd_text_right(x_right, y, text, size=10, bold=False, color=TEXT):
    return cmd_text(x_right - text_width(text, size, bold), y, text, size, bold, color)


def cmd_text_center(center_x, y, text, size=10, bold=False, color=TEXT):
    return cmd_text(center_x - text_width(text, size, bold) / 2, y, text, size, bold, color)


def cmd_line(x1, y1, x2, y2, width=0.7, color=BORDER):
    return (
        f"{rgb_cmd(color, stroke=True)} {width:.3f} w "
        f"{x1:.2f} {y1:.2f} m {x2:.2f} {y2:.2f} l S"
    )


def cmd_rect(x, y, w, h, width=0.7, fill=None, stroke=BORDER):
    parts = []
    if fill is not None:
        parts.append(f"{rgb_cmd(fill)} {x:.2f} {y:.2f} {w:.2f} {h:.2f} re f")
    if stroke is not None and width > 0:
        parts.append(
            f"{rgb_cmd(stroke, stroke=True)} {width:.3f} w "
            f"{x:.2f} {y:.2f} {w:.2f} {h:.2f} re S"
        )
    return "\n".join(parts)


def cmd_round_rect(x, y, w, h, radius=5, fill=None, stroke=None, width=0.8):
    # PDF cubic Bezier approximation of a rounded rectangle.
    r = min(radius, w / 2, h / 2)
    k = 0.5522847498307936
    p = [
        f"{x+r:.2f} {y:.2f} m",
        f"{x+w-r:.2f} {y:.2f} l",
        f"{x+w-r+k*r:.2f} {y:.2f} {x+w:.2f} {y+r-k*r:.2f} {x+w:.2f} {y+r:.2f} c",
        f"{x+w:.2f} {y+h-r:.2f} l",
        f"{x+w:.2f} {y+h-r+k*r:.2f} {x+w-r+k*r:.2f} {y+h:.2f} {x+w-r:.2f} {y+h:.2f} c",
        f"{x+r:.2f} {y+h:.2f} l",
        f"{x+r-k*r:.2f} {y+h:.2f} {x:.2f} {y+h-r+k*r:.2f} {x:.2f} {y+h-r:.2f} c",
        f"{x:.2f} {y+r:.2f} l",
        f"{x:.2f} {y+r-k*r:.2f} {x+r-k*r:.2f} {y:.2f} {x+r:.2f} {y:.2f} c",
        "h",
    ]
    prefix = []
    op = "S"
    if fill is not None:
        prefix.append(rgb_cmd(fill))
        op = "f"
    if stroke is not None:
        prefix.append(rgb_cmd(stroke, stroke=True))
        prefix.append(f"{width:.3f} w")
        op = "B" if fill is not None else "S"
    return " ".join(prefix + p + [op])


# Top-origin helpers. Approved PDFs were measured in a top-left coordinate system.
PAGE_H = 792.0


def ty(top_y):
    return PAGE_H - top_y


def t_text(x, baseline_top, text, size=10, bold=False, color=TEXT):
    return cmd_text(x, ty(baseline_top), text, size, bold, color)


def t_text_right(x_right, baseline_top, text, size=10, bold=False, color=TEXT):
    return cmd_text_right(x_right, ty(baseline_top), text, size, bold, color)


def t_text_center(center_x, baseline_top, text, size=10, bold=False, color=TEXT):
    return cmd_text_center(center_x, ty(baseline_top), text, size, bold, color)


def t_line(x1, y1_top, x2, y2_top, width=0.7, color=BORDER):
    return cmd_line(x1, ty(y1_top), x2, ty(y2_top), width, color)


def t_rect(x, top_y, w, h, width=0.7, fill=None, stroke=BORDER):
    return cmd_rect(x, PAGE_H - top_y - h, w, h, width, fill, stroke)


def t_round_rect(x, top_y, w, h, radius=5, fill=None, stroke=None, width=0.8):
    return cmd_round_rect(
        x, PAGE_H - top_y - h, w, h, radius, fill, stroke, width
    )


def fit_text(text, max_width, size=9, bold=False, min_size=None):
    text = str(text or "")
    if min_size is not None:
        s = size
        while s > min_size and text_width(text, s, bold) > max_width:
            s -= 0.2
        if text_width(text, s, bold) <= max_width:
            return text, s
    if text_width(text, size, bold) <= max_width:
        return text
    t = text
    while t and text_width(t + "...", size, bold) > max_width:
        t = t[:-1]
    return t.rstrip() + "..."


def wrap_text(text, max_width, size=9, bold=False):
    text = str(text or "").replace("\xa0", " ")
    all_lines = []
    for paragraph in text.splitlines() or [""]:
        words = paragraph.split()
        if not words:
            continue
        cur = words[0]
        for word in words[1:]:
            trial = cur + " " + word
            if text_width(trial, size, bold) <= max_width:
                cur = trial
            else:
                all_lines.append(cur)
                cur = word
        all_lines.append(cur)
    return all_lines


def draw_t_multiline(cmds, x, first_baseline_top, lines, size=9, leading=11,
                     bold=False, color=TEXT, max_lines=None):
    if isinstance(lines, str):
        lines = lines.splitlines()
    lines = list(lines)
    if max_lines is not None:
        lines = lines[:max_lines]
    y = first_baseline_top
    for line in lines:
        cmds.append(t_text(x, y, line, size, bold, color))
        y += leading
    return y


def combined_row_description(row):
    """Preserve useful QuickBooks descriptions without repeating generic text.

    QuickBooks often exports the same description twice in different forms:
      Item Description: "Packing & Handling Charges"
      Memo/Description: "Packing & Handling Charges Tracking # ..."

    Keep the richer version and suppress the shorter text when it is already
    fully contained inside another description field.
    """
    candidates = []

    for name in ("Memo", "Item Description", "Description", "Line Description"):
        value = get(row, name)
        if not value:
            continue

        value = value.strip()
        normalized = re.sub(r"\s+", " ", value).strip().lower()
        if not normalized:
            continue

        candidates.append((value, normalized))

    # First remove exact duplicates while preserving the original field order.
    unique = []
    seen = set()
    for value, normalized in candidates:
        if normalized in seen:
            continue
        seen.add(normalized)
        unique.append((value, normalized))

    # Then remove a shorter/generic description if its complete normalized text
    # is already contained in a richer description. This prevents output like:
    #   Packing & Handling Charges
    #   Tracking # ...
    #   Packing & Handling Charges
    kept = []
    for i, (value, normalized) in enumerate(unique):
        contained_in_richer = False
        for j, (_other_value, other_norm) in enumerate(unique):
            if i == j:
                continue
            if len(other_norm) > len(normalized) and normalized in other_norm:
                contained_in_richer = True
                break
        if not contained_in_richer:
            kept.append(value)

    return "\n".join(kept)


def format_tracking_override_for_pdf(value):
    """Format a manual tracking entry without doubling Tracking/FedEx labels.

    Examples:
      8778 6322 0018
        -> Tracking # 8778 6322 0018

      Tracking #: 8778 6322 0018
        -> Tracking #: 8778 6322 0018

      Fedex # 8778 6322 0018
        -> Fedex # 8778 6322 0018

      Tracking # Tracking #: 8778 6322 0018
        -> Tracking #: 8778 6322 0018
    """
    raw = str(value or "").strip()
    if not raw:
        return ""

    # Collapse accidental repeated tracking labels from a pasted/manual value.
    raw = re.sub(
        r"(?i)^\s*tracking\s*#?\s*:?\s*tracking\s*#?\s*:?\s*",
        "Tracking #: ",
        raw,
        count=1
    )

    # If the user already supplied a meaningful label, preserve it.
    if re.match(r"(?i)^\s*(tracking|fed\s*ex|fedex)\b", raw):
        return raw

    return f"Tracking # {raw}"


def invoice_item_rows(inv: Invoice):
    out = []
    for r in inv.rows:
        if is_header_row(r):
            continue
        acct = get(r, "Account").lower()
        if acct == "sales tax payable":
            continue
        item = get(r, "Item")
        if not item:
            continue

        qty = abs(dec(get(r, "Qty")) or ZERO)
        rate = abs(dec(get(r, "Sales Price")) or ZERO)
        amt = abs(dec(get(r, "Amount")) or ZERO)
        memo = get(r, "Memo")
        generic = get(r, "Item Description")

        clean_item = re.sub(r"\s+\([^)]*\)\s*$", "", item).strip()
        desc = combined_row_description(r) or memo or generic

        if "packing" in clean_item.lower() and "handling" in clean_item.lower():
            desc = re.sub(
                r"^(Packing\s*&\s*Handling\s*Charges)\s+",
                r"\1\n",
                desc,
                flags=re.IGNORECASE
            )
            desc = re.sub(
                r"^(Packing\s*&\s*Handling\s*Charges)\n?(?=\d{4}\s+\d{4}\s+\d{4})",
                r"\1\n",
                desc,
                flags=re.IGNORECASE
            )
            if inv.tracking_override:
                desc = (desc or "Packing & Handling Charges").rstrip()
                override_digits = extract_tracking_number(inv.tracking_override)
                existing_digits = extract_tracking_number(desc)
                if override_digits and override_digits != existing_digits:
                    desc += f"\n{format_tracking_override_for_pdf(inv.tracking_override)}"
            elif inv.tracking_detected_text and not looks_like_tracking(desc):
                desc = (desc or "Packing & Handling Charges").rstrip()
                desc += f"\n{inv.tracking_detected_text.strip()}"

        out.append({
            "item": clean_item,
            "description": desc,
            "qty": qty,
            "uom": "ea" if "label" in clean_item.lower() else "",
            "rate": rate,
            "amount": amt,
        })
    return out


def _wrapped_limited(text, max_width, size, max_lines, bold=False):
    lines = wrap_text(text, max_width, size, bold)
    if len(lines) <= max_lines:
        return lines
    clipped = lines[:max_lines]
    last = clipped[-1]
    while last and text_width(last + "...", size, bold) > max_width:
        last = last[:-1]
    clipped[-1] = last.rstrip() + "..."
    return clipped


def add_invoice_page(pdf: SimplePDF, inv: Invoice, cfg):
    c = []

    # Header
    c.append(t_text(42, 54, cfg["company_name"], 15, True, NAVY))
    for line, base in zip(cfg["company_address_lines"], (72, 84, 96)):
        c.append(t_text(42, base, line, 8.5, False, MUTED))

    c.append(t_text_right(570, 60, f"Invoice #{inv.num}", 20, True, TEXT))
    c.append(t_text_right(570, 84, inv.date, 9, False, MUTED))
    c.append(t_line(42, 112, 570, 112, 1.4, NAVY))

    # Bill To / Ship To cards
    cards = [
        (42, "BILL TO", cfg["bill_to_lines"]),
        (313, "SHIP TO", inv.ship_to_block.splitlines() if inv.ship_to_block else ["SHIP TO REVIEW REQUIRED"]),
    ]
    for x, title, lines in cards:
        c.append(t_round_rect(x, 132, 257, 108, 5, None, BORDER, 0.8))
        c.append(t_round_rect(x, 132, 257, 22, 5, LIGHT_GRAY, None, 0))
        c.append(t_rect(x, 148, 257, 6, 0, LIGHT_GRAY, None))
        c.append(t_text(x + 10, 147, title, 7.5, True, MUTED))
        for idx, line in enumerate(lines[:7]):
            c.append(t_text(
                x + 10, 170 + idx * 11, line,
                8.6, idx == 0, TEXT
            ))

    # Terms strip
    meta = [
        (42, 92, "TERMS", inv.terms),
        (134, 92, "S.O. NO.", inv.so_num),
        (226, 150, "SHIP VIA", inv.via),
        (376, 194, "PO NO.", inv.po),
    ]
    for x, w, label, value in meta:
        c.append(t_rect(x, 256, w, 48, 0.7, LIGHT_GRAY, BORDER))
        c.append(t_text(x + 8, 270, label, 6.8, True, MUTED))
        fitted = fit_text(value, w - 16, 8.5, True, min_size=7.0)
        if isinstance(fitted, tuple):
            txt, sz = fitted
        else:
            txt, sz = fitted, 8.5
        c.append(t_text(x + 8, 287, txt, sz, True, TEXT))

    # Item table
    xs = [42, 140, 338, 384, 440, 476, 526, 570]
    widths = [xs[i+1] - xs[i] for i in range(len(xs)-1)]
    headers = ["Item", "Description", "Prev.", "Qty", "U/M", "Rate", "Amount"]
    c.append(t_rect(42, 322, 528, 24, 0, NAVY, None))
    for i, label in enumerate(headers):
        if label == "Amount":
            c.append(t_text_right(xs[i+1] - 8, 338, label, 7.4, True, WHITE))
        else:
            c.append(t_text(xs[i] + 7, 338, label, 7.4, True, WHITE))

    rows = invoice_item_rows(inv)
    row_top = 346.0
    for row in rows:
        item_lines = _wrapped_limited(row["item"], widths[0] - 14, 7.8, 2, True)
        desc_lines = _wrapped_limited(row["description"], widths[1] - 14, 8.0, 4, False)
        max_lines = max(len(item_lines), len(desc_lines), 1)
        row_h = 20 + 10 * max_lines
        if max_lines >= 3:
            row_h += 1

        c.append(t_rect(42, row_top, 528, row_h, 0.6, None, BORDER))
        for vx in xs[1:-1]:
            c.append(t_line(vx, row_top, vx, row_top + row_h, 0.6, BORDER))

        for idx, line in enumerate(item_lines):
            c.append(t_text(49, row_top + 15 + idx * 10, line, 7.8, True, TEXT))
        for idx, line in enumerate(desc_lines):
            c.append(t_text(147, row_top + 15 + idx * 10, line, 8.0, False, TEXT))

        numeric_base = row_top + (19 if row_h <= 40 else 19)
        prev_x0, prev_x1 = xs[2], xs[3]
        qty_x0, qty_x1 = xs[3], xs[4]
        uom_x0, uom_x1 = xs[4], xs[5]
        rate_x0, rate_x1 = xs[5], xs[6]
        amt_x0, amt_x1 = xs[6], xs[7]

        c.append(t_text_center((prev_x0 + prev_x1) / 2, numeric_base, "0", 8, False, TEXT))
        c.append(t_text_center((qty_x0 + qty_x1) / 2, numeric_base, f"{row['qty']:,.0f}", 8, False, TEXT))
        c.append(t_text_center((uom_x0 + uom_x1) / 2, numeric_base, row["uom"], 8, False, TEXT))
        rate_s = f"{row['rate']:.2f}" if row["rate"] >= Decimal("1") else f"{row['rate']:.5f}".rstrip("0").rstrip(".")
        c.append(t_text_center((rate_x0 + rate_x1) / 2, numeric_base, rate_s, 8, False, TEXT))
        c.append(t_text_right(amt_x1 - 8, numeric_base, f"{row['amount']:,.2f}", 8.2, True, TEXT))

        row_top += row_h

    # Message and totals. Coworkers sometimes leave the QuickBooks customer
    # message blank, so every generated invoice gets the standard thank-you
    # message when no source message is available.
    summary_base = row_top + 28
    invoice_message = (inv.customer_message or "").strip() or DEFAULT_INVOICE_MESSAGE
    message_lines = _wrapped_limited(invoice_message, 310, 8.5, 3, True)
    for idx, line in enumerate(message_lines):
        c.append(t_text(
            42, summary_base + idx * 11,
            line,
            8.5, True, TEXT
        ))

    item_subtotal = sum((r["amount"] for r in rows), ZERO)
    tax_amount = inv.original_total - item_subtotal
    if abs(tax_amount) < Decimal("0.005"):
        tax_amount = ZERO

    totals = [
        ("Subtotal", item_subtotal, False, False),
        (cfg["tax_label"], tax_amount, False, False),
        ("Total", inv.original_total, True, True),
        ("Payments/Credits", inv.payments_credits, False, False),
        ("Balance Due", inv.open_balance, True, True),
    ]

    top = summary_base - 10
    for idx, (label, value, bold, shaded) in enumerate(totals):
        row_t = top + idx * 22
        if shaded:
            c.append(t_round_rect(380, row_t, 190, 22, 3, LIGHT_BLUE, None, 0))
        if idx > 0 and not shaded:
            c.append(t_line(380, row_t, 570, row_t, 0.5, BORDER))
        if idx > 0 and shaded:
            c.append(t_line(380, row_t, 570, row_t, 0.5, BORDER))
        c.append(t_text(388, row_t + 16, label, 8.2, bold, TEXT))
        c.append(t_text_right(562, row_t + 16, money(value), 8.4, bold, TEXT))

    # Footer
    c.append(t_line(42, 716, 570, 716, 0.7, BORDER))
    c.append(t_text(42, 736, cfg["company_phone"], 7.8, False, MUTED))
    c.append(t_text_center(306, 736, cfg["company_email"], 7.8, False, MUTED))
    c.append(t_text_right(570, 736, cfg["company_website"], 7.8, False, MUTED))

    pdf.add_page(c)


def generate_invoice_packet(result: ValidationResult, cfg, path: Path):
    pdf = SimplePDF()
    for n in result.include_nums:
        add_invoice_page(pdf, result.invoices[n], cfg)
    pdf.save(path)


def _balance_period_title(period):
    if not period:
        return "Billing Period"
    return f"{period[0]:%B} {period[0].day} - {period[1].day}, {period[1].year}"


def add_balance_page(pdf, result, cfg, rows_subset, page_no, total_pages,
                     show_total=False):
    c = []

    c.append(t_text_center(306, 42, cfg["company_name"], 14, True, NAVY))
    c.append(t_text_center(306, 63, "Customer Open Balance", 20, True, TEXT))
    c.append(t_text_center(306, 80, _balance_period_title(result.period), 8.5, True, MUTED))
    c.append(t_line(42, 94, 570, 94, 1.2, NAVY))

    c.append(t_round_rect(42, 108, 528, 21, 4, LIGHT_BLUE, None, 0))
    c.append(t_text(52, 122, cfg["customer_name"], 8.2, True, NAVY))

    c.append(t_rect(42, 136, 528, 21, 1.2, LIGHT_GRAY, BORDER))
    c.append(t_text(50, 150, "Type", 7.1, True, TEXT))
    c.append(t_text(118, 150, "Date", 7.1, True, TEXT))
    c.append(t_text(200, 150, "Invoice #", 7.1, True, TEXT))
    c.append(t_text(270, 150, "Due Date", 7.1, True, TEXT))
    c.append(t_text_right(464, 150, "Open Balance", 7.1, True, TEXT))
    c.append(t_text_right(562, 150, "Amount", 7.1, True, TEXT))

    baseline = 166.5
    row_h = 13.3
    for idx, n in enumerate(rows_subset):
        inv = result.invoices[n]
        bal = result.balances[n]
        c.append(t_text(50, baseline, "Invoice", 6.9, False, TEXT))
        c.append(t_text(118, baseline, inv.date, 6.9, False, TEXT))
        c.append(t_text(200, baseline, inv.num, 6.9, True, TEXT))
        c.append(t_text(270, baseline, str(bal["due_date"]), 6.9, False, TEXT))
        c.append(t_text_right(464, baseline, money(inv.open_balance), 6.9, False, TEXT))
        c.append(t_text_right(562, baseline, money(inv.balance_report_amount), 6.9, False, TEXT))

        line_top = 170.3 + idx * row_h
        c.append(t_line(42, line_top, 570, line_top, 0.3, ROW_LINE))
        baseline += row_h

    last_line_top = 157 + len(rows_subset) * row_h

    if show_total:
        divider_top = last_line_top + 10
        total_top = divider_top + 14
        c.append(t_line(42, divider_top, 570, divider_top, 1.0, NAVY))
        c.append(t_round_rect(42, total_top, 528, 24, 4, NAVY, None, 0))
        c.append(t_text(52, total_top + 16, "TOTAL", 8.7, True, WHITE))
        c.append(t_text_right(464, total_top + 16, money(result.open_total), 8.7, True, WHITE))
        c.append(t_text_right(562, total_top + 16, money(result.original_total), 8.7, True, WHITE))

    c.append(t_text_center(306, 772, f"Page {page_no}", 6.8, False, MUTED))
    pdf.add_page(c)


def generate_open_balance(result: ValidationResult, cfg, path: Path):
    pdf = SimplePDF()

    # 35 rows exactly reproduce the approved one-page test layout.
    # Larger batches paginate while preserving the same visual system.
    per_page = 35
    chunks = [
        result.include_nums[i:i + per_page]
        for i in range(0, len(result.include_nums), per_page)
    ] or [[]]

    for i, chunk in enumerate(chunks, 1):
        add_balance_page(
            pdf, result, cfg, chunk, i, len(chunks),
            show_total=(i == len(chunks))
        )
    pdf.save(path)

def format_period(period):
    if not period:
        return "Unknown"
    return f"{period[0]:%m/%d/%Y} - {period[1]:%m/%d/%Y}"


def period_filename(period):
    if not period:
        return datetime.now().strftime("%m-%d-%Y")
    return f"{period[0]:%m-%d-%Y} to {period[1]:%m-%d-%Y}"


def load_history():
    try:
        v = json.loads(HISTORY_PATH.read_text(encoding="utf-8"))
        return v if isinstance(v, list) else []
    except Exception:
        return []


def save_history(rows):
    HISTORY_PATH.write_text(json.dumps(rows, indent=2), encoding="utf-8")


class BillingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        ensure_files()
        self.cfg = load_config()
        self.ship_store = ShipToStore()
        self.session_ship = {}
        self.result: Optional[ValidationResult] = None
        self.generated_paths = []
        self.tracking_review_pending = False
        self.outlook_draft_opened = False
        self.attachment_reminder = None
        self.title(APP_NAME)
        # Fit the available desktop, including laptops using Windows DPI scaling.
        # Never require a width larger than the usable screen.
        screen_w, screen_h = self.winfo_screenwidth(), self.winfo_screenheight()
        window_w = min(1180, max(680, screen_w - 80))
        window_h = min(760, max(460, screen_h - 100))
        window_w = min(window_w, screen_w)
        window_h = min(window_h, screen_h)
        self.geometry(f"{window_w}x{window_h}")
        self.minsize(min(780, screen_w), min(520, screen_h))
        self.configure(bg=COLORS["light"])
        self._build_styles()
        self._build_ui()

        # Keep keyboard focus support, but don't leave a dotted focus rectangle
        # sitting on the last button the user clicked with the mouse.
        self.bind_class(
            "TButton",
            "<ButtonRelease-1>",
            self._clear_button_focus_after_mouse,
            add="+"
        )

    def _clear_button_focus_after_mouse(self, event):
        """Clear mouse-click focus after the button command finishes.

        This removes the persistent dotted focus outline without disabling
        keyboard focus/navigation for users who Tab through the interface.
        """
        try:
            top = event.widget.winfo_toplevel()
        except Exception:
            top = self

        def clear_focus():
            try:
                if top.winfo_exists():
                    top.focus_set()
            except Exception:
                pass

        self.after_idle(clear_focus)

    def _build_styles(self):
        s = ttk.Style(self)
        try:
            s.theme_use("clam")
        except Exception:
            pass

        s.configure("TFrame", background=COLORS["light"])
        s.configure("Card.TFrame", background=COLORS["card"])

        s.configure(
            "Title.TLabel",
            background=COLORS["light"],
            foreground=COLORS["navy"],
            font=("Segoe UI", 20, "bold")
        )
        s.configure(
            "Subtitle.TLabel",
            background=COLORS["light"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 10)
        )
        s.configure(
            "Period.TLabel",
            background=COLORS["light"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 10)
        )
        s.configure(
            "DialogMeta.TLabel",
            background=COLORS["light"],
            foreground=COLORS["text"],
            font=("Segoe UI", 10, "bold")
        )
        s.configure(
            "Section.TLabel",
            background=COLORS["card"],
            foreground=COLORS["navy"],
            font=("Segoe UI", 10, "bold")
        )

        # Summary cards
        s.configure(
            "CardValue.TLabel",
            background=COLORS["card"],
            foreground=COLORS["navy"],
            font=("Segoe UI", 17, "bold")
        )
        s.configure(
            "CardValueMuted.TLabel",
            background=COLORS["card"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 17, "bold")
        )
        s.configure(
            "CardValueGood.TLabel",
            background=COLORS["card"],
            foreground=COLORS["green"],
            font=("Segoe UI", 14, "bold")
        )
        s.configure(
            "CardValueWarn.TLabel",
            background=COLORS["card"],
            foreground=COLORS["amber"],
            font=("Segoe UI", 13, "bold")
        )
        s.configure(
            "CardValueBad.TLabel",
            background=COLORS["card"],
            foreground=COLORS["red"],
            font=("Segoe UI", 13, "bold")
        )
        s.configure(
            "CardLabel.TLabel",
            background=COLORS["card"],
            foreground=COLORS["muted"],
            font=("Segoe UI", 8)
        )

        s.configure(
            "File.TEntry",
            padding=(8, 8),
            font=("Segoe UI", 9),
            fieldbackground="#FFFFFF"
        )

        # Buttons intentionally have visible depth again: raised relief, a light
        # top edge, darker lower edge, and a distinct pressed state.
        s.configure(
            "Browse.TButton",
            font=("Segoe UI", 9),
            padding=(12, 8),
            background="#E7EAEE",
            foreground=COLORS["text"],
            borderwidth=2,
            relief="raised",
            bordercolor="#AAB2BC",
            lightcolor="#FFFFFF",
            darkcolor="#AAB3BD"
        )
        s.map(
            "Browse.TButton",
            background=[
                ("active", "#F0F2F5"),
                ("pressed", "#D8DDE3"),
            ],
            relief=[
                ("pressed", "sunken"),
                ("!pressed", "raised"),
            ]
        )

        # Current next action: dimensional navy button.
        s.configure(
            "Primary.TButton",
            font=("Segoe UI", 9, "bold"),
            padding=(11, 9),
            background=COLORS["navy"],
            foreground="#FFFFFF",
            borderwidth=2,
            relief="raised",
            bordercolor="#0E263B",
            lightcolor="#315574",
            darkcolor="#0B2236"
        )
        s.map(
            "Primary.TButton",
            background=[
                ("active", "#244967"),
                ("pressed", "#102A42"),
                ("disabled", "#E4E7EB"),
            ],
            foreground=[
                ("disabled", "#A2A8B0"),
            ],
            bordercolor=[
                ("disabled", "#D8DDE3"),
            ],
            lightcolor=[
                ("disabled", "#F4F5F7"),
            ],
            darkcolor=[
                ("disabled", "#CED3D9"),
            ],
            relief=[
                ("pressed", "sunken"),
                ("!pressed", "raised"),
            ]
        )

        # Completed/available actions: neutral raised buttons.
        s.configure(
            "Neutral.TButton",
            # Match Primary.TButton font metrics exactly so switching the
            # highlighted/current step never changes button geometry.
            font=("Segoe UI", 9, "bold"),
            padding=(11, 9),
            background="#E8EBEF",
            foreground=COLORS["navy"],
            borderwidth=2,
            relief="raised",
            bordercolor="#AAB2BC",
            lightcolor="#FFFFFF",
            darkcolor="#A9B2BC"
        )
        s.map(
            "Neutral.TButton",
            background=[
                ("active", "#F0F2F5"),
                ("pressed", "#D6DCE2"),
                ("disabled", "#F0F2F4"),
            ],
            foreground=[
                ("disabled", "#A4A9B0"),
            ],
            bordercolor=[
                ("disabled", "#D9DEE4"),
            ],
            lightcolor=[
                ("disabled", "#F8F9FA"),
            ],
            darkcolor=[
                ("disabled", "#D8DDE2"),
            ],
            relief=[
                ("pressed", "sunken"),
                ("!pressed", "raised"),
            ]
        )

        # Smaller utility action, still visibly button-like.
        s.configure(
            "Utility.TButton",
            font=("Segoe UI", 8),
            padding=(9, 7),
            background="#ECEFF2",
            foreground=COLORS["text"],
            borderwidth=2,
            relief="raised",
            bordercolor="#B1B8C1",
            lightcolor="#FFFFFF",
            darkcolor="#B1B9C2"
        )
        s.map(
            "Utility.TButton",
            background=[
                ("active", "#F4F5F7"),
                ("pressed", "#DCE1E6"),
            ],
            relief=[
                ("pressed", "sunken"),
                ("!pressed", "raised"),
            ]
        )

        # Popup secondary actions share the same dimensional neutral treatment.
        s.configure(
            "Secondary.TButton",
            font=("Segoe UI", 9),
            padding=(10, 8),
            background="#E8EBEF",
            foreground=COLORS["navy"],
            borderwidth=2,
            relief="raised",
            bordercolor="#AAB2BC",
            lightcolor="#FFFFFF",
            darkcolor="#A9B2BC"
        )
        s.map(
            "Secondary.TButton",
            background=[
                ("active", "#F0F2F5"),
                ("pressed", "#D6DCE2"),
                ("disabled", "#F0F2F4"),
            ],
            foreground=[
                ("disabled", "#A4A9B0"),
            ],
            relief=[
                ("pressed", "sunken"),
                ("!pressed", "raised"),
            ]
        )

    def _build_ui(self):
        outer = ttk.Frame(self, padding=10)
        outer.pack(fill="both", expand=True)

        # Pack the action bar FIRST along the bottom. The expanding progress
        # panel can no longer displace buttons below the window edge.
        self.actions = ttk.Frame(outer)
        self.actions.pack(side="bottom", fill="x", pady=(4, 0))
        self.actions.columnconfigure(1, weight=1)

        # Header
        header = ttk.Frame(outer)
        header.pack(fill="x", pady=(2, 8))
        ttk.Label(header, text="BILLING RECONCILIATION DEMO", style="Title.TLabel").pack(anchor="center")
        self.period_var = tk.StringVar(value="Billing Period: Not validated")
        ttk.Label(header, textvariable=self.period_var, style="Period.TLabel").pack(
            anchor="center", pady=(4, 0)
        )

        # QuickBooks exports. Full path stays internal; only filename is shown.
        self.detail_path = ""
        self.balance_path = ""
        self.detail_display_var = tk.StringVar(value="No file selected")
        self.balance_display_var = tk.StringVar(value="No file selected")
        self.file_buttons = {}

        file_card = ttk.Frame(outer, style="Card.TFrame", padding=14)
        file_card.pack(fill="x", pady=(12, 10))
        ttk.Label(file_card, text="QUICKBOOKS EXPORTS", style="Section.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 9)
        )

        self._file_row(
            file_card, 1, "Invoice Data", "detail",
            self.detail_display_var, self.choose_detail
        )
        self._file_row(
            file_card, 2, "Open Balance", "balance",
            self.balance_display_var, self.choose_balance
        )
        file_card.columnconfigure(1, weight=1)

        # Summary cards
        summary_box = ttk.Frame(outer)
        summary_box.pack(fill="x", pady=7)

        self.card_vars = {}
        self.card_value_labels = {}
        specs = [
            ("found", "Invoices Found"),
            ("included", "Invoices Included"),
            ("balance", "Balance Due"),
            ("payments", "Payments / Credits"),
            ("issues", None),
        ]

        for i, (key, label) in enumerate(specs):
            f = ttk.Frame(summary_box, style="Card.TFrame", padding=(12, 10))
            f.grid(row=0, column=i, padx=(0 if i == 0 else 5, 0), sticky="nsew")
            summary_box.columnconfigure(i, weight=1)

            var = tk.StringVar(value="--")
            self.card_vars[key] = var
            value_label = ttk.Label(f, textvariable=var, style="CardValue.TLabel")
            self.card_value_labels[key] = value_label

            if key == "issues":
                # The value explains itself ("✓ No issues", "1 error", etc.),
                # so this card intentionally has no redundant caption.
                value_label.pack(expand=True)
            else:
                value_label.pack()
                ttk.Label(f, text=label, style="CardLabel.TLabel").pack()

        # BILLING PROGRESS is a true scrollable widget area.
        # Progress lines, bold section headings, and the live status banner all
        # live inside the same scrollable region.
        # Outer Billing Progress panel and inner console intentionally use
        # different backgrounds so the console reads as its own workspace.
        self.review_card = tk.Frame(
            outer,
            bg="#EEF2F6",
            bd=0,
            padx=14,
            pady=12
        )
        self.review_card.pack(fill="both", expand=True, pady=(10, 10))

        tk.Label(
            self.review_card,
            text="BILLING PROGRESS",
            bg="#EEF2F6",
            fg=COLORS["navy"],
            font=("Segoe UI", 10, "bold")
        ).pack(anchor="w")

        progress_wrap = tk.Frame(
            self.review_card,
            bg="#FFFFFF",
            highlightbackground="#D7DEE8",
            highlightthickness=1,
            bd=0
        )
        progress_wrap.pack(fill="both", expand=True, pady=(8, 0))
        progress_wrap.rowconfigure(0, weight=1)
        progress_wrap.columnconfigure(0, weight=1)

        self.progress_canvas = tk.Canvas(
            progress_wrap,
            highlightthickness=0,
            borderwidth=0,
            background="#FFFFFF"
        )
        self.progress_scrollbar = ttk.Scrollbar(
            progress_wrap,
            orient="vertical",
            command=self.progress_canvas.yview
        )
        self.progress_canvas.configure(
            yscrollcommand=self.progress_scrollbar.set
        )

        self.progress_canvas.grid(row=0, column=0, sticky="nsew")
        self.progress_scrollbar.grid(row=0, column=1, sticky="ns")
        self.progress_scrollbar.grid_remove()

        self.progress_inner = tk.Frame(
            self.progress_canvas,
            bg="#FFFFFF"
        )
        self._progress_window_id = self.progress_canvas.create_window(
            (0, 0),
            window=self.progress_inner,
            anchor="nw"
        )

        self.progress_inner.bind(
            "<Configure>",
            lambda e: self.progress_canvas.configure(
                scrollregion=self.progress_canvas.bbox("all")
            )
        )
        self.progress_canvas.bind(
            "<Configure>",
            self._resize_progress_inner
        )
        self.progress_canvas.bind("<Enter>", self._progress_mouse_enter)
        self.progress_canvas.bind("<Leave>", self._progress_mouse_leave)
        self.progress_inner.bind("<Enter>", self._progress_mouse_enter)
        self.progress_inner.bind("<Leave>", self._progress_mouse_leave)

        self._progress_widgets = []
        self._last_progress_label = None

        # One live status banner. It sits at the bottom of the scrollable
        # Billing Progress area and is moved below new progress automatically.
        self.progress_status_frame = tk.Frame(
            self.progress_inner,
            bg="#FFFFFF"
        )
        self.progress_status_label = tk.Label(
            self.progress_status_frame,
            justify="left",
            anchor="w",
            font=("Segoe UI", 10, "bold"),
            padx=14,
            pady=9,
            wraplength=1000,
            bg="#FFFFFF",
            fg="#FBFCFD",
            text=""
        )
        self.progress_status_label.pack(fill="x", expand=True)

        # The desktop action bar always stays in ONE horizontal row.
        # Smaller widths reduce button size and abbreviate long action labels;
        # the original four-action order and right alignment are preserved.
        self.manage_btn = ttk.Button(
            self.actions, text="Manage Ship Tos", style="Utility.TButton",
            width=16, command=self.manage_ship_tos,
        )
        self.manage_btn.grid(row=0, column=0, sticky="w", padx=(0, 10))

        self.workflow = ttk.Frame(self.actions)
        self.workflow.grid(row=0, column=1, sticky="e")
        workflow_width = 24  # original desktop sizing when space allows
        self.validate_btn = ttk.Button(
            self.workflow, text="VALIDATE BILLING", style="Primary.TButton",
            width=workflow_width, command=self.validate_action,
        )
        self.generate_btn = ttk.Button(
            self.workflow, text="GENERATE BILLING PDFs", style="Neutral.TButton",
            width=workflow_width, command=self.generate_pdfs,
        )
        self.open_btn = ttk.Button(
            self.workflow, text="REVIEW PDFs", style="Neutral.TButton",
            width=workflow_width, command=self.open_pdfs,
        )
        self.outlook_btn = ttk.Button(
            self.workflow, text="CREATE OUTLOOK DRAFT", style="Neutral.TButton",
            width=workflow_width, state="disabled", command=self.create_outlook_draft,
        )
        self._toolbar_size = None
        self._toolbar_layout_busy = False
        self._layout_action_buttons()
        self.actions.bind("<Configure>", self._layout_action_buttons)
        self._set_button_state("selected")

    def _layout_action_buttons(self, event=None):
        """Fit four equal-width workflow buttons into a single horizontal row.

        Candidate sizes are checked against the actual ttk geometry. This
        accommodates different Windows font metrics and display scaling.
        """
        if self._toolbar_layout_busy:
            return
        width = self.actions.winfo_width()
        if width <= 1:
            width = max(0, self.winfo_width() - 20)
        self._toolbar_layout_busy = True
        try:
            buttons = (
                self.validate_btn, self.generate_btn, self.open_btn, self.outlook_btn,
            )
            captions = (
                ("VALIDATE BILLING", "GENERATE BILLING PDFs", "REVIEW PDFs", "CREATE OUTLOOK DRAFT"),
                ("VALIDATE BILLING", "GENERATE PDFs", "REVIEW PDFs", "OUTLOOK DRAFT"),
                ("VALIDATE", "GENERATE", "REVIEW", "OUTLOOK"),
            )
            choices = (
                (24, 0),  # original desktop widths and full labels
                (22, 0),
                (20, 0),  # 1024px laptop normally fits here
                (18, 1),
                (16, 1),
                (14, 1),
                (12, 2),  # small window or elevated display scaling
            )
            selected = None
            for chars, label_mode in choices:
                for button, caption in zip(buttons, captions[label_mode]):
                    button.configure(width=chars, text=caption)
                self.update_idletasks()
                required = (
                    self.manage_btn.winfo_reqwidth() + 10
                    + sum(button.winfo_reqwidth() for button in buttons)
                    + 3 * 8
                )
                selected = (chars, label_mode)
                if required <= width - 2:
                    break
            self._toolbar_size = selected
            for i, button in enumerate(buttons):
                button.grid(row=0, column=i, padx=(0 if i == 0 else 8, 0))
        finally:
            self._toolbar_layout_busy = False

    def _file_row(self, parent, row, label, key, var, command):
        ttk.Label(parent, text=label, style="Section.TLabel").grid(
            row=row, column=0, sticky="w", padx=(0, 12), pady=6
        )
        ent = ttk.Entry(
            parent,
            textvariable=var,
            style="File.TEntry",
            state="readonly"
        )
        ent.grid(row=row, column=1, sticky="ew", pady=6)

        btn = ttk.Button(
            parent,
            text="Browse...",
            command=command,
            style="Browse.TButton"
        )
        btn.grid(row=row, column=2, padx=(9, 0), pady=6, sticky="ns")
        self.file_buttons[key] = btn

    def _select_file(self, key, path):
        path = str(path)
        filename = Path(path).name
        display = f"✓ {filename}"

        if key == "detail":
            self.detail_path = path
            self.detail_display_var.set(display)
        else:
            self.balance_path = path
            self.balance_display_var.set(display)

        self.file_buttons[key].configure(text="Change...")
        self._reset_after_file_change()

    def choose_detail(self):
        p = filedialog.askopenfilename(
            title="Select Invoice Data CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if p:
            self._select_file("detail", p)

    def choose_balance(self):
        p = filedialog.askopenfilename(
            title="Select Open Balance CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if p:
            self._select_file("balance", p)

    def _reset_after_file_change(self):
        if self.attachment_reminder is not None:
            try:
                if self.attachment_reminder.winfo_exists():
                    self.attachment_reminder.destroy()
            except Exception:
                pass
            self.attachment_reminder = None

        self.result = None
        self.generated_paths = []
        self.tracking_review_pending = False
        self.outlook_draft_opened = False
        self.period_var.set("Billing Period: Not validated")
        self.reset_cards()
        self._hide_banner()
        self._clear_review()
        self._set_button_state("selected")

    def _set_button_state(self, stage):
        self.manage_btn.configure(state="normal", style="Utility.TButton")

        # Start from a neutral workflow appearance.
        self.validate_btn.configure(style="Neutral.TButton")
        self.generate_btn.configure(style="Neutral.TButton")
        self.open_btn.configure(style="Neutral.TButton")
        self.outlook_btn.configure(style="Neutral.TButton")

        # Validate stays available except while it is actively running.
        self.validate_btn.configure(state="disabled" if stage == "validating" else "normal")

        if stage == "ready":
            self.generate_btn.configure(state="normal", style="Primary.TButton")
            self.open_btn.configure(state="disabled")
            self.outlook_btn.configure(state="disabled")

        elif stage == "pdfs":
            self.generate_btn.configure(state="normal")
            self.open_btn.configure(state="normal", style="Primary.TButton")
            self.outlook_btn.configure(state="disabled")

        elif stage == "reviewed":
            self.generate_btn.configure(state="normal")
            self.open_btn.configure(state="normal")
            self.outlook_btn.configure(state="normal", style="Primary.TButton")

        elif stage == "outlook":
            self.generate_btn.configure(state="normal")
            self.open_btn.configure(state="normal")
            self.outlook_btn.configure(state="normal")

        elif stage in {"selected", "needs_review"}:
            self.validate_btn.configure(style="Primary.TButton")
            self.generate_btn.configure(state="disabled")
            self.open_btn.configure(state="disabled")
            self.outlook_btn.configure(state="disabled")

        elif stage == "validating":
            self.generate_btn.configure(state="disabled")
            self.open_btn.configure(state="disabled")
            self.outlook_btn.configure(state="disabled")

        else:
            self.generate_btn.configure(state="disabled")
            self.open_btn.configure(state="disabled")
            self.outlook_btn.configure(state="disabled")

    def _resize_progress_inner(self, event):
        self.progress_canvas.itemconfigure(
            self._progress_window_id,
            width=event.width
        )
        self.progress_status_label.configure(
            wraplength=max(500, event.width - 50)
        )

    def _on_progress_mousewheel(self, event):
        if not self.progress_canvas.bbox("all"):
            return "break"
        steps = int(-1 * (event.delta / 120)) if event.delta else 0
        if steps:
            self.progress_canvas.yview_scroll(steps, "units")
        return "break"

    def _progress_mouse_enter(self, _event):
        self.bind_all("<MouseWheel>", self._on_progress_mousewheel)

    def _progress_mouse_leave(self, _event):
        self.unbind_all("<MouseWheel>")

    def _scroll_progress_bottom(self):
        """Keep latest work visible and show the scrollbar only when needed."""
        self.update_idletasks()
        bbox = self.progress_canvas.bbox("all")
        if not bbox:
            self.progress_scrollbar.grid_remove()
            self.progress_canvas.yview_moveto(0.0)
            return

        self.progress_canvas.configure(scrollregion=bbox)
        content_height = max(0, bbox[3] - bbox[1])
        viewport_height = max(1, self.progress_canvas.winfo_height())
        needs_scroll = content_height > (viewport_height + 2)

        if needs_scroll:
            self.progress_scrollbar.grid()
            self.progress_canvas.yview_moveto(1.0)
        else:
            self.progress_scrollbar.grid_remove()
            self.progress_canvas.yview_moveto(0.0)

    def _status_before_widget(self):
        return (
            self.progress_status_frame
            if self.progress_status_frame.winfo_manager()
            else None
        )

    def _show_banner(self, text, kind="good"):
        palette = {
            "good": (COLORS["green_bg"], COLORS["green"]),
            "warn": (COLORS["amber_bg"], COLORS["amber"]),
            "bad": (COLORS["red_bg"], COLORS["red"]),
        }
        bg, fg = palette.get(kind, palette["good"])

        self.progress_status_frame.configure(bg=bg)
        self.progress_status_label.configure(
            bg=bg,
            fg=fg,
            text=text
        )

        if not self.progress_status_frame.winfo_manager():
            self.progress_status_frame.pack(
                fill="x",
                padx=0,
                pady=(12, 4)
            )
        else:
            # Re-pack so the one status strip always sits below the newest line.
            self.progress_status_frame.pack_forget()
            self.progress_status_frame.pack(
                fill="x",
                padx=0,
                pady=(12, 4)
            )

        self._scroll_progress_bottom()

    def _hide_banner(self):
        if self.progress_status_frame.winfo_manager():
            self.progress_status_frame.pack_forget()
        self.progress_status_frame.configure(bg="#FFFFFF")
        self.progress_status_label.configure(
            bg="#FFFFFF",
            fg="#FBFCFD",
            text=""
        )

    def set_status(self, text, kind="warn"):
        pass

    def _clear_review(self):
        for widget in list(self._progress_widgets):
            try:
                widget.destroy()
            except Exception:
                pass
        self._progress_widgets.clear()
        self._last_progress_label = None
        self._hide_banner()
        self.progress_canvas.yview_moveto(0.0)

    def _pack_progress_widget(self, widget, *, pady=(0, 0), fill="x"):
        before = self._status_before_widget()
        kwargs = dict(
            side="top",
            anchor="w",
            fill=fill,
            padx=10,
            pady=pady
        )
        if before is not None:
            kwargs["before"] = before
        widget.pack(**kwargs)
        self._progress_widgets.append(widget)
        self._scroll_progress_bottom()

    def _append_review(self, line="", pause=0.0, tag="normal"):
        if tag == "file":
            font = ("Segoe UI", 9)
            fg = COLORS["muted"]
            padx = 20
        else:
            font = ("Segoe UI", 10)
            fg = COLORS["text"]
            padx = 10

        label = tk.Label(
            self.progress_inner,
            text=line,
            justify="left",
            anchor="w",
            font=font,
            bg="#FFFFFF",
            fg=fg,
            wraplength=1000
        )

        before = self._status_before_widget()
        kwargs = dict(
            side="top",
            anchor="w",
            fill="x",
            padx=(padx, 10),
            pady=(1, 1)
        )
        if before is not None:
            kwargs["before"] = before
        label.pack(**kwargs)

        self._progress_widgets.append(label)
        self._last_progress_label = label
        self._scroll_progress_bottom()

        if pause:
            self.update()
            time.sleep(pause)

    def _review_section(self, title):
        label = tk.Label(
            self.progress_inner,
            text=title,
            justify="left",
            anchor="w",
            font=("Segoe UI", 10, "bold"),
            bg="#FFFFFF",
            fg=COLORS["navy"]
        )

        before = self._status_before_widget()
        kwargs = dict(
            side="top",
            anchor="w",
            fill="x",
            padx=10,
            pady=(12 if self._progress_widgets else 2, 4)
        )
        if before is not None:
            kwargs["before"] = before
        label.pack(**kwargs)

        self._progress_widgets.append(label)
        self._last_progress_label = None
        self._scroll_progress_bottom()

    def _replace_last_review_line(self, text, tag="normal"):
        if self._last_progress_label is None:
            self._append_review(text, tag=tag)
            return

        self._last_progress_label.configure(text=text)
        self._scroll_progress_bottom()
        self.update_idletasks()
        self.update()

    def _replace_progress_line_starting(self, prefix, text):
        for widget in self._progress_widgets:
            try:
                if str(widget.cget("text")).startswith(prefix):
                    widget.configure(text=text)
                    self._scroll_progress_bottom()
                    self.update_idletasks()
                    self.update()
                    return True
            except Exception:
                pass
        return False

    def _refresh_issue_section(self, result):
        """Replace the current issue section without replaying validation."""
        start = None
        for i, widget in enumerate(self._progress_widgets):
            try:
                if str(widget.cget("text")) == "ISSUES REQUIRING REVIEW":
                    start = i
                    break
            except Exception:
                pass

        if start is not None:
            # Remove the spacer immediately before the issue heading too.
            if start > 0:
                try:
                    if str(self._progress_widgets[start - 1].cget("text")) == "":
                        start -= 1
                except Exception:
                    pass

            for widget in self._progress_widgets[start:]:
                try:
                    widget.destroy()
                except Exception:
                    pass
            self._progress_widgets = self._progress_widgets[:start]
            self._last_progress_label = None

        self._append_issues(result)

    def _same_validation_files(self, detail_path, balance_path):
        if not self.result:
            return False
        try:
            return (
                Path(self.result.detail_path).resolve() == Path(detail_path).resolve()
                and Path(self.result.balance_path).resolve() == Path(balance_path).resolve()
            )
        except Exception:
            return (
                str(self.result.detail_path) == str(detail_path)
                and str(self.result.balance_path) == str(balance_path)
            )

    def _tracking_pending_banner_text(self, result):
        found = len(result.invoices)
        included = len(result.include_nums)
        remaining = len(self.tracking_review_items(result))

        if found == included:
            base = f"{included} invoices reconciled."
        else:
            excluded = found - included
            base = (
                f"{found} invoices found. {included} included. "
                f"{excluded} paid/closed excluded."
            )

        noun = "tracking number is" if remaining == 1 else "tracking numbers are"
        return (
            "READY WITH WARNINGS\n"
            f"{base} {remaining} {noun} still missing. "
            "To add the missing tracking, click Validate Billing again."
        )

    def _retry_pending_tracking(self):
        """Open only Tracking Review for a previously accepted warning set."""
        result = self.result
        if not result:
            return

        self._set_button_state("validating")
        self.validate_btn.configure(state="disabled")

        resolved = self.resolve_tracking_warnings(result)
        self.tracking_review_pending = (
            not resolved and bool(self.tracking_review_items(result))
        )

        # Any newly entered tracking changes the PDFs, so old generated copies
        # must not remain the active review targets.
        if resolved:
            self.generated_paths = []
            self.outlook_draft_opened = False

        self._set_summary_cards(result)
        self._refresh_issue_section(result)

        if self.tracking_review_pending:
            self._show_banner(
                self._tracking_pending_banner_text(result),
                "warn"
            )
        else:
            self._show_banner(
                self._ready_banner_text(result),
                "warn" if result.warnings else "good"
            )

        self.validate_btn.configure(state="normal")
        self._set_button_state("ready")

    def set_review(self, text):
        self._clear_review()
        for line in str(text or "").splitlines():
            self._append_review(line)

    def reset_cards(self):
        for key, var in self.card_vars.items():
            var.set("--")
            self.card_value_labels[key].configure(style="CardValue.TLabel")

    def _set_summary_cards(self, result):
        hard = len(result.hard_errors)
        warnings = len(result.warnings)

        self.card_vars["found"].set(str(len(result.invoices)))
        self.card_vars["included"].set(str(len(result.include_nums)))
        self.card_vars["balance"].set(money(result.open_total))
        self.card_vars["payments"].set(money(result.payments_total))

        if hard == 0 and warnings == 0:
            self.card_vars["issues"].set("✓ No issues")
        else:
            parts = []
            if hard:
                parts.append(f"{hard} error{'s' if hard != 1 else ''}")
            if warnings:
                parts.append(f"{warnings} warning{'s' if warnings != 1 else ''}")
            self.card_vars["issues"].set(" · ".join(parts))

        if result.payments_total == ZERO:
            self.card_value_labels["payments"].configure(style="CardValueMuted.TLabel")
        else:
            self.card_value_labels["payments"].configure(style="CardValue.TLabel")

        if hard:
            self.card_value_labels["issues"].configure(style="CardValueBad.TLabel")
        elif warnings:
            self.card_value_labels["issues"].configure(style="CardValueWarn.TLabel")
        else:
            self.card_value_labels["issues"].configure(style="CardValueGood.TLabel")

    def _issue_codes(self, result, level=None):
        issues = result.issues
        if level:
            issues = [i for i in issues if i.level == level]
        return {i.code for i in issues}

    def _animate_one_check(self, working_text, finished_text, pause=0.85):
        self._append_review(working_text)
        time.sleep(pause)
        self._replace_last_review_line(finished_text)
        time.sleep(0.28)

    def _animate_validation_steps(self, result):
        """Show each validation stage clearly, one at a time."""
        hard_codes = self._issue_codes(result, "HARD")
        warn_codes = self._issue_codes(result, "WARNING")

        packing_hard = {
            "WRONG_PACKING_ITEM",
            "WRONG_PACKING_RATE",
            "WRONG_PACKING_AMOUNT",
        }
        packing_warn = {
            "PACKING_LINE_MISSING",
            "MULTIPLE_PACKING_LINES",
        }
        tracking_warn = {"TRACKING_NOT_FOUND"}
        reconcile_hard = hard_codes - packing_hard

        reconcile_result = (
            "✕ Invoice reports need review"
            if reconcile_hard
            else "✓ Invoice reports reconciled"
        )
        self._animate_one_check(
            "Reconciling Invoice Data with Open Balance...",
            reconcile_result
        )

        if hard_codes & packing_hard:
            packing_result = "✕ Packing & Handling needs correction"
        elif warn_codes & packing_warn:
            packing_result = "⚠ Packing & Handling needs review"
        else:
            packing_result = "✓ Packing & Handling checked"
        self._animate_one_check(
            "Checking Packing & Handling...",
            packing_result
        )

        tracking_result = (
            "⚠ Tracking needs review"
            if warn_codes & tracking_warn
            else "✓ Tracking checked"
        )
        self._animate_one_check(
            "Checking tracking numbers...",
            tracking_result
        )

    def _append_issues(self, result):
        if not result.hard_errors and not result.warnings:
            return
        self._append_review("")
        self._append_review("ISSUES REQUIRING REVIEW")
        for issue in result.hard_errors:
            who = f"Invoice {issue.invoice}: " if issue.invoice else ""
            self._append_review(f"HARD ERROR | {who}{issue.message}")
        for issue in result.warnings:
            if issue.code == "TRACKING_NOT_FOUND" and issue.invoice in result.invoices:
                inv = result.invoices[issue.invoice]
                po = inv.po or "[blank]"
                self._append_review(
                    f"WARNING | PO {po} | Invoice {issue.invoice}: Missing tracking number."
                )
            else:
                who = f"Invoice {issue.invoice}: " if issue.invoice else ""
                self._append_review(f"WARNING | {who}{issue.message}")

    def _ready_banner_text(self, result):
        found = len(result.invoices)
        included = len(result.include_nums)
        warnings = len(result.warnings)

        if found == included:
            base = f"{included} invoices reconciled."
        else:
            excluded = found - included
            base = f"{found} invoices found. {included} included. {excluded} paid/closed excluded."

        if warnings:
            suffix = f" {warnings} warning{'s' if warnings != 1 else ''} require review."
            return "READY WITH WARNINGS\n" + base + suffix
        return "READY TO GENERATE BILLING PDFs\n" + base + " No issues require review."

    def validate_action(self):
        dp = Path(self.detail_path) if self.detail_path else None
        bp = Path(self.balance_path) if self.balance_path else None

        if not dp or not bp or not dp.exists() or not bp.exists():
            messagebox.showerror(
                "Select both CSV files",
                "Choose the Invoice Data CSV and the Open Balance CSV first."
            )
            return

        # If the user previously chose Continue With Warnings for missing
        # tracking, Validate Billing becomes a shortcut back to Tracking Review.
        # The completed validation animation does not replay.
        if (
            self.tracking_review_pending
            and self._same_validation_files(dp, bp)
            and self.tracking_review_items(self.result)
        ):
            self._retry_pending_tracking()
            return

        self.tracking_review_pending = False
        self.generated_paths = []
        self.reset_cards()
        self._set_button_state("validating")
        self.validate_btn.configure(state="disabled")
        self._clear_review()

        try:
            result = validate_files(dp, bp, self.cfg)
            self.result = result
            self.period_var.set(f"Billing Period: {format_period(result.period)}")

            # The first visible line is always this real bold subheading.
            self._review_section("BILLING VALIDATION")
            self._animate_validation_steps(result)

            if result.hard_errors:
                self._set_summary_cards(result)
                self._append_issues(result)
                self._show_banner(
                    "✕ Billing requires review.\n"
                    "Next: Correct the hard error(s) in QuickBooks, re-export both reports, then validate again.",
                    "bad"
                )
                self.validate_btn.configure(state="normal")
                self._set_button_state("needs_review")
                messagebox.showwarning(
                    "Billing requires review",
                    "Hard validation errors were found.\n\n"
                    "Correct them in QuickBooks, re-export BOTH reports for the same week, then validate again."
                )
                return

            # Ship To is the fourth Billing Validation check.
            self._append_review("Verifying Ship To addresses...")
            time.sleep(0.85)
            self._replace_last_review_line("Checking Ship To addresses...")

            if not self.resolve_ship_tos(result):
                self._replace_last_review_line("⚠ Ship To review not finished")
                self._set_summary_cards(result)
                self._append_issues(result)
                self._show_banner(
                    "⚠ Ship To review is not finished.\n"
                    "Next: Finish the Ship To review, then validate billing again.",
                    "warn"
                )
                self.validate_btn.configure(state="normal")
                self._set_button_state("needs_review")
                return

            self.apply_ship_blocks(result)
            self._replace_last_review_line("✓ Ship Tos verified")
            time.sleep(0.28)

            # Missing tracking can be resolved inside the bot. If the user
            # chooses Continue With Warnings, remember that state so the next
            # Validate Billing click can reopen Tracking Review immediately.
            tracking_resolved = self.resolve_tracking_warnings(result)
            self.tracking_review_pending = (
                not tracking_resolved and bool(self.tracking_review_items(result))
            )

            # Populate the cards only after all review steps complete.
            self._set_summary_cards(result)

            excluded = len(result.invoices) - len(result.include_nums)
            if excluded > 0:
                self._append_review(
                    f"{excluded} paid/closed "
                    f"{'invoice' if excluded == 1 else 'invoices'} excluded from this billing run."
                )

            self._append_issues(result)

            if self.tracking_review_pending:
                self._show_banner(
                    self._tracking_pending_banner_text(result),
                    "warn"
                )
            else:
                self._show_banner(
                    self._ready_banner_text(result),
                    "warn" if result.warnings else "good"
                )
            self.validate_btn.configure(state="normal")
            self._set_button_state("ready")

        except Exception as e:
            technical_log("ERROR\n" + traceback.format_exc())
            if self.result:
                self._set_summary_cards(self.result)
            self._review_section("ERROR")
            self._append_review(str(e))
            self._show_banner(
                "✕ Validation could not be completed.\n"
                "Next: Resolve the error above, then validate billing again.",
                "bad"
            )
            self.validate_btn.configure(state="normal")
            self._set_button_state("needs_review")
            messagebox.showerror(
                APP_NAME,
                f"{e}\n\nTechnical details were saved in the logs folder."
            )

    def ship_clue_text(self, inv: Invoice):
        raw = inv.ship_raw
        vals = []
        preferred = [
            "Ship To Address 1__2", "Ship To Address 2",
            "Ship To City", "Ship To State", "Ship Zip", "Ship To Address 1"
        ]
        for key in preferred:
            val = raw.get(key, "")
            if val:
                vals.append(val)
        # De-duplicate while preserving order.
        seen = set()
        cleaned = []
        for v in vals:
            k = v.strip().lower()
            if k and k not in seen:
                seen.add(k)
                cleaned.append(v.strip())
        return "\n".join(cleaned) or "(QuickBooks report did not provide readable Ship To clues.)"

    def resolve_ship_tos(self, result):
        by_sig = defaultdict(list)
        for n in result.include_nums:
            inv = result.invoices[n]
            if not inv.ship_signature:
                messagebox.showerror("Ship To missing",
                                     f"Invoice {n} has no Ship To signature in the Invoice Data CSV.")
                return False
            by_sig[inv.ship_signature].append(n)

        unknown = []
        for sig, nums in by_sig.items():
            if self.ship_store.get(sig) or self.session_ship.get(sig):
                continue
            inv = result.invoices[nums[0]]
            historical_match = self.ship_store.find_unique_match(inv.ship_raw)
            if historical_match:
                self.ship_store.set(
                    sig,
                    historical_match,
                    source="Automatically matched to a verified historical Ship To"
                )
                continue
            unknown.append((sig, nums))

        if not unknown:
            return True

        # Make the console state explicit before the modal review begins.
        self._replace_last_review_line("⚠ Ship To addresses need review")
        self.update_idletasks()
        self.update()

        for idx, (sig, nums) in enumerate(unknown, 1):
            inv = result.invoices[nums[0]]
            block = self.ask_ship_to(inv, nums, idx, len(unknown))
            if block is None:
                return False
            text, remember = block
            if remember:
                self.ship_store.set(sig, text)
            else:
                self.session_ship[sig] = text
        return True

    def ask_ship_to(self, inv, nums, index, total):
        win = tk.Toplevel(self)
        win.title("Ship To Review Required")
        win.geometry("760x680")
        win.minsize(660, 580)
        win.transient(self)
        win.grab_set()
        win.configure(bg=COLORS["light"])
        result = {"value": None}

        outer = ttk.Frame(win, padding=20)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(5, weight=1)

        ttk.Label(
            outer,
            text=f"Ship To Review Required   {index} of {total}",
            style="Title.TLabel"
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            outer,
            text=f"Invoice{'s' if len(nums) > 1 else ''}: {', '.join(nums)}",
            style="DialogMeta.TLabel"
        ).grid(row=1, column=0, sticky="w", pady=(4, 14))

        clues = ttk.Frame(outer, style="Card.TFrame", padding=12)
        clues.grid(row=2, column=0, sticky="ew")
        ttk.Label(clues, text="QUICKBOOKS CLUES", style="Section.TLabel").pack(anchor="w")
        ttk.Label(
            clues,
            text=self.ship_clue_text(inv),
            background=COLORS["card"],
            foreground=COLORS["text"],
            font=("Segoe UI", 9),
            justify="left"
        ).pack(anchor="w", pady=(7, 0))

        instruction_invoice = nums[0]
        instruction = tk.Label(
            outer,
            text=(
                f"Open invoice {instruction_invoice} in QuickBooks and copy the entire Ship To box exactly as shown, "
                "or choose a saved Ship To:"
            ),
            foreground=COLORS["navy"],
            background=COLORS["light"],
            font=("Segoe UI", 10, "bold"),
            anchor="w",
            justify="left",
            wraplength=690,
            padx=6,
            pady=5
        )
        instruction.grid(row=3, column=0, sticky="ew", pady=(16, 6))

        pulse_colors = ("#FFF0B8", COLORS["light"])
        def pulse_instruction(step=0):
            # Three quick highlight pulses, then remain still.
            if not win.winfo_exists() or step >= 6:
                if win.winfo_exists():
                    instruction.configure(background=COLORS["light"])
                return
            instruction.configure(background=pulse_colors[step % 2])
            win.after(260, lambda: pulse_instruction(step + 1))

        win.after(250, pulse_instruction)

        text = scrolledtext.ScrolledText(
            outer,
            height=8,
            wrap="word",
            font=("Segoe UI", 10)
        )
        text.grid(row=4, column=0, sticky="nsew")

        def choose_saved_ship_to():
            self.ship_store.load()

            picker = tk.Toplevel(win)
            picker.title("Choose Saved Ship To")
            picker.geometry("900x570")
            picker.minsize(740, 480)
            picker.transient(win)
            picker.grab_set()
            picker.configure(bg=COLORS["light"])

            shell = ttk.Frame(picker, padding=18)
            shell.pack(fill="both", expand=True)
            shell.columnconfigure(0, weight=1)
            shell.rowconfigure(3, weight=1)

            ttk.Label(
                shell,
                text="Choose Saved Ship To",
                style="Title.TLabel"
            ).grid(row=0, column=0, sticky="w")
            ttk.Label(
                shell,
                text=(
                    "Select the verified address that should print on these invoices. "
                    "Likely matches to the QuickBooks clues are shown first."
                ),
                style="Subtitle.TLabel",
                wraplength=840,
                justify="left"
            ).grid(row=1, column=0, sticky="ew", pady=(3, 12))

            search_row = ttk.Frame(shell)
            search_row.grid(row=2, column=0, sticky="ew", pady=(0, 10))
            search_row.columnconfigure(1, weight=1)
            ttk.Label(
                search_row,
                text="Search:",
                style="DialogMeta.TLabel"
            ).grid(row=0, column=0, sticky="w", padx=(0, 8))
            search_var = tk.StringVar()
            search_entry = ttk.Entry(
                search_row,
                textvariable=search_var,
                style="File.TEntry"
            )
            search_entry.grid(row=0, column=1, sticky="ew")

            body = ttk.Frame(shell)
            body.grid(row=3, column=0, sticky="nsew")
            body.columnconfigure(0, weight=1)
            body.columnconfigure(1, weight=1)
            body.rowconfigure(0, weight=1)

            left = ttk.Frame(body, style="Card.TFrame", padding=10)
            left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
            right = ttk.Frame(body, style="Card.TFrame", padding=10)
            right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

            ttk.Label(
                left,
                text="SAVED SHIP TOS",
                style="Section.TLabel"
            ).pack(anchor="w", pady=(0, 6))

            tree = ttk.Treeview(
                left,
                columns=("name", "detail"),
                show="headings",
                height=15
            )
            tree.heading("name", text="Name")
            tree.heading("detail", text="Location / Detail")
            tree.column("name", width=190, anchor="w")
            tree.column("detail", width=220, anchor="w")
            tree.pack(fill="both", expand=True)

            ttk.Label(
                right,
                text="FULL ADDRESS BLOCK",
                style="Section.TLabel"
            ).pack(anchor="w", pady=(0, 6))
            preview = scrolledtext.ScrolledText(
                right,
                height=15,
                wrap="word",
                font=("Segoe UI", 10),
                background="#FFFFFF",
                foreground=COLORS["text"],
                padx=10,
                pady=10
            )
            preview.pack(fill="both", expand=True)
            preview.configure(state="disabled")

            # Build search/sort tokens from the current QuickBooks clues.
            clue_text = self.ship_clue_text(inv).lower()
            clue_tokens = {
                t for t in re.findall(r"[a-z0-9]+", clue_text)
                if len(t) >= 3 and t not in {
                    "attn", "tel", "phone", "usa", "ext", "the", "and"
                }
            }

            records = []
            seen_blocks = set()
            for key, rec in self.ship_store.all_unique_blocks():
                block = (rec.get("address_block") or "").strip()
                if not block:
                    continue
                norm_block = self.ship_store._norm(block)
                if norm_block in seen_blocks:
                    continue
                seen_blocks.add(norm_block)

                lines = [x.strip() for x in block.splitlines() if x.strip()]
                name = rec.get("label") or (lines[0] if lines else "Saved Ship To")
                detail = lines[1] if len(lines) > 1 else ""
                block_tokens = set(re.findall(r"[a-z0-9]+", block.lower()))
                score = len(clue_tokens & block_tokens)

                records.append({
                    "key": key,
                    "name": name,
                    "detail": detail,
                    "block": block,
                    "source": rec.get("source", ""),
                    "score": score,
                })

            records.sort(
                key=lambda r: (-r["score"], r["name"].lower(), r["block"].lower())
            )

            record_by_iid = {}

            def refresh(*_):
                tree.delete(*tree.get_children())
                record_by_iid.clear()
                q = search_var.get().strip().lower()
                for rec in records:
                    hay = (rec["name"] + "\n" + rec["detail"] + "\n" + rec["block"]).lower()
                    if q and q not in hay:
                        continue
                    iid = tree.insert(
                        "",
                        "end",
                        values=(rec["name"], rec["detail"])
                    )
                    record_by_iid[iid] = rec

                children = tree.get_children()
                if children:
                    tree.selection_set(children[0])
                    tree.focus(children[0])
                    show_selected()

            def selected_record():
                sel = tree.selection()
                return record_by_iid.get(sel[0]) if sel else None

            def show_selected(_event=None):
                rec = selected_record()
                if not rec:
                    return
                display = rec["block"]
                if rec["source"]:
                    display += "\n\nSource: " + rec["source"]
                preview.configure(state="normal")
                preview.delete("1.0", "end")
                preview.insert("1.0", display)
                preview.configure(state="disabled")

            def use_selected():
                rec = selected_record()
                if not rec:
                    messagebox.showinfo(
                        "Select a Ship To",
                        "Select a saved address first.",
                        parent=picker
                    )
                    return
                text.delete("1.0", "end")
                text.insert("1.0", rec["block"])
                picker.destroy()
                text.focus_set()

            tree.bind("<<TreeviewSelect>>", show_selected)
            tree.bind("<Double-1>", lambda _event: use_selected())
            search_var.trace_add("write", refresh)

            actions = ttk.Frame(shell)
            actions.grid(row=4, column=0, sticky="ew", pady=(10, 0))
            actions.columnconfigure(1, weight=1)
            ttk.Button(
                actions,
                text="Cancel",
                command=picker.destroy,
                style="Secondary.TButton"
            ).grid(row=0, column=0, sticky="w")
            ttk.Button(
                actions,
                text="USE SELECTED SHIP TO",
                command=use_selected,
                style="Primary.TButton"
            ).grid(row=0, column=2, sticky="e")

            refresh()
            search_entry.focus_set()

        remember_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            outer,
            text="Remember this Ship To for future invoices",
            variable=remember_var
        ).grid(row=5, column=0, sticky="w", pady=(10, 10))

        btn = ttk.Frame(outer)
        btn.grid(row=6, column=0, sticky="ew")
        btn.columnconfigure(0, weight=1)
        btn.columnconfigure(1, weight=1)
        btn.columnconfigure(2, weight=1)

        def cancel():
            result["value"] = None
            win.destroy()

        def save():
            v = text.get("1.0", "end").strip()
            lines = [x.strip() for x in v.splitlines() if x.strip()]
            if len(lines) < 2:
                messagebox.showerror(
                    "Complete Ship To",
                    "Paste the complete multi-line Ship To block from QuickBooks or choose a saved Ship To.",
                    parent=win
                )
                return
            result["value"] = ("\n".join(lines), remember_var.get())
            win.destroy()

        ttk.Button(
            btn,
            text="Cancel Billing",
            command=cancel,
            style="Secondary.TButton"
        ).grid(row=0, column=0, sticky="w")

        ttk.Button(
            btn,
            text="CHOOSE SAVED SHIP TO...",
            command=choose_saved_ship_to,
            style="Secondary.TButton"
        ).grid(row=0, column=1)

        ttk.Button(
            btn,
            text="SAVE & CONTINUE",
            command=save,
            style="Primary.TButton"
        ).grid(row=0, column=2, sticky="e")

        win.protocol("WM_DELETE_WINDOW", cancel)
        self.wait_window(win)
        return result["value"]

    def tracking_review_items(self, result):
        seen = set()
        items = []
        for issue in result.warnings:
            if issue.code != "TRACKING_NOT_FOUND" or not issue.invoice:
                continue
            if issue.invoice in seen or issue.invoice not in result.invoices:
                continue
            seen.add(issue.invoice)
            items.append(result.invoices[issue.invoice])
        return items

    def ask_tracking_numbers(self, result):
        items = self.tracking_review_items(result)
        if not items:
            return {"mode": "none", "values": {}}

        win = tk.Toplevel(self)
        win.title("Tracking Review Required")
        win.geometry("900x560")
        win.minsize(760, 460)
        win.transient(self)
        win.grab_set()
        win.configure(bg=COLORS["light"])
        outcome = {"value": None}

        outer = ttk.Frame(win, padding=20)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(3, weight=1)

        ttk.Label(
            outer,
            text=f"Tracking Review Required   {len(items)} invoice{'s' if len(items) != 1 else ''}",
            style="Title.TLabel"
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            outer,
            text=(
                "Search the customer's order email thread using the PO number, then locate and paste the tracking number below. "
                "These entries are used in the generated billing PDFs only; the bot does not change QuickBooks."
            ),
            style="Subtitle.TLabel",
            wraplength=820,
            justify="left"
        ).grid(row=1, column=0, sticky="ew", pady=(4, 14))

        header = tk.Frame(outer, bg="#E8EDF3", padx=10, pady=7)
        header.grid(row=2, column=0, sticky="ew")
        header.columnconfigure(0, weight=4)
        header.columnconfigure(1, weight=2)
        header.columnconfigure(2, weight=5)
        tk.Label(header, text="PO", bg="#E8EDF3", fg=COLORS["navy"],
                 font=("Segoe UI", 9, "bold"), anchor="w").grid(row=0, column=0, sticky="ew")
        tk.Label(header, text="Invoice", bg="#E8EDF3", fg=COLORS["navy"],
                 font=("Segoe UI", 9, "bold"), anchor="w").grid(row=0, column=1, sticky="ew")
        tk.Label(header, text="Tracking Number", bg="#E8EDF3", fg=COLORS["navy"],
                 font=("Segoe UI", 9, "bold"), anchor="w").grid(row=0, column=2, sticky="ew")

        body_wrap = tk.Frame(outer, bg="#FFFFFF", highlightbackground=COLORS["border"],
                             highlightthickness=1, bd=0)
        body_wrap.grid(row=3, column=0, sticky="nsew")
        body_wrap.rowconfigure(0, weight=1)
        body_wrap.columnconfigure(0, weight=1)

        canvas = tk.Canvas(body_wrap, background="#FFFFFF", highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(body_wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        rows = tk.Frame(canvas, bg="#FFFFFF")
        window_id = canvas.create_window((0, 0), window=rows, anchor="nw")
        rows.columnconfigure(0, weight=4)
        rows.columnconfigure(1, weight=2)
        rows.columnconfigure(2, weight=5)

        entry_vars = {}
        for i, inv in enumerate(items):
            bg = "#FFFFFF" if i % 2 == 0 else "#FAFBFC"
            tk.Label(rows, text=inv.po or "[blank]", bg=bg, fg=COLORS["text"],
                     font=("Segoe UI", 9, "bold"), anchor="w", padx=10, pady=10,
                     wraplength=280, justify="left").grid(row=i, column=0, sticky="nsew")
            tk.Label(rows, text=inv.num, bg=bg, fg=COLORS["text"],
                     font=("Segoe UI", 9), anchor="w", padx=10, pady=10).grid(row=i, column=1, sticky="nsew")
            cell = tk.Frame(rows, bg=bg, padx=8, pady=7)
            cell.grid(row=i, column=2, sticky="nsew")
            var = tk.StringVar()
            entry_vars[inv.num] = var
            ttk.Entry(cell, textvariable=var, style="File.TEntry").pack(fill="x")

        def sync_scrollregion(_event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfigure(window_id, width=canvas.winfo_width())
            need = rows.winfo_reqheight() > canvas.winfo_height() + 2
            if need:
                scrollbar.grid()
            else:
                scrollbar.grid_remove()
                canvas.yview_moveto(0.0)

        rows.bind("<Configure>", sync_scrollregion)
        canvas.bind("<Configure>", sync_scrollregion)

        buttons = ttk.Frame(outer)
        buttons.grid(row=4, column=0, sticky="ew", pady=(10, 0))
        buttons.columnconfigure(1, weight=1)

        def continue_with_warnings():
            outcome["value"] = {"mode": "warnings", "values": {}}
            win.destroy()

        def apply_values():
            values = {}
            bad = []
            for inv in items:
                raw = entry_vars[inv.num].get().strip()
                if not raw or not looks_like_tracking(raw):
                    bad.append(f"PO {inv.po or '[blank]'} / Invoice {inv.num}")
                else:
                    values[inv.num] = raw
            if bad:
                messagebox.showerror(
                    "Tracking number required",
                    "Enter a valid tracking number for each row, or choose Continue With Warnings.\n\n"
                    + "Needs attention:\n" + "\n".join(bad),
                    parent=win
                )
                return
            outcome["value"] = {"mode": "apply", "values": values}
            win.destroy()

        ttk.Button(
            buttons, text="Continue With Warnings", command=continue_with_warnings,
            style="Secondary.TButton"
        ).grid(row=0, column=0, sticky="w")
        ttk.Button(
            buttons, text="APPLY TRACKING & CONTINUE", command=apply_values,
            style="Primary.TButton"
        ).grid(row=0, column=2, sticky="e")

        win.protocol("WM_DELETE_WINDOW", continue_with_warnings)
        self.wait_window(win)
        return outcome["value"] or {"mode": "warnings", "values": {}}

    def resolve_tracking_warnings(self, result):
        items = self.tracking_review_items(result)
        if not items:
            return True

        review = self.ask_tracking_numbers(result)
        if review.get("mode") != "apply":
            return False

        values = review.get("values", {})
        for inv_num, tracking in values.items():
            if inv_num in result.invoices:
                result.invoices[inv_num].tracking_override = tracking.strip()

        resolved = set(values)
        result.issues = [
            issue for issue in result.issues
            if not (issue.code == "TRACKING_NOT_FOUND" and issue.invoice in resolved)
        ]

        count = len(resolved)
        self._replace_progress_line_starting(
            "⚠ Tracking needs review",
            f"✓ Tracking reviewed ({count} added in bot)"
        )
        technical_log(
            "Tracking overrides added for "
            + ", ".join(
                f"Invoice {n} / PO {result.invoices[n].po}: {result.invoices[n].tracking_override}"
                for n in sorted(resolved)
            )
        )
        return True

    def apply_ship_blocks(self, result):
        for n in result.include_nums:
            inv = result.invoices[n]
            inv.ship_to_block = self.session_ship.get(inv.ship_signature) or self.ship_store.get(inv.ship_signature)

    def manage_ship_tos(self):
        self.ship_store.load()
        win = tk.Toplevel(self)
        win.title("Manage Ship To Addresses")
        win.geometry("980x660")
        win.minsize(760, 560)
        win.transient(self)

        outer = ttk.Frame(win, padding=16)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Manage Ship To Addresses", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            outer,
            text="Search and preview the full multi-line Ship To blocks known to the billing bot.",
            style="Subtitle.TLabel"
        ).pack(anchor="w", pady=(2, 10))

        search_var = tk.StringVar()
        search_row = ttk.Frame(outer)
        search_row.pack(fill="x", pady=(0, 10))
        ttk.Label(search_row, text="Search:", style="Section.TLabel").pack(side="left", padx=(0, 8))
        search_entry = ttk.Entry(search_row, textvariable=search_var, style="File.TEntry")
        search_entry.pack(side="left", fill="x", expand=True)

        body = ttk.Frame(outer)
        body.pack(fill="both", expand=True)
        body.columnconfigure(0, weight=1)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        left_frame = ttk.Frame(body, style="Card.TFrame", padding=10)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 7))
        right_frame = ttk.Frame(body, style="Card.TFrame", padding=10)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(7, 0))

        ttk.Label(left_frame, text="KNOWN SHIP TOS", style="Section.TLabel").pack(anchor="w", pady=(0, 6))
        tree = ttk.Treeview(left_frame, columns=("label",), show="headings", height=14)
        tree.heading("label", text="Ship To")
        tree.column("label", width=390)
        tree.pack(fill="both", expand=True)

        ttk.Label(right_frame, text="FULL ADDRESS BLOCK", style="Section.TLabel").pack(anchor="w", pady=(0, 6))
        preview = scrolledtext.ScrolledText(
            right_frame, height=14, wrap="word", font=("Segoe UI", 10),
            background="#FFFFFF", foreground=COLORS["text"], padx=10, pady=10
        )
        preview.pack(fill="both", expand=True)
        preview.configure(state="disabled")

        key_by_item = {}
        records = []

        def rebuild_records():
            nonlocal records
            records = []
            seen_blocks = set()
            for key, rec in self.ship_store.all_unique_blocks():
                block = (rec.get("address_block") or "").strip()
                if not block:
                    continue
                nk = self.ship_store._norm(block)
                if nk in seen_blocks:
                    continue
                seen_blocks.add(nk)
                records.append({
                    "key": key,
                    "label": rec.get("label") or block.splitlines()[0],
                    "block": block,
                    "source": rec.get("source", ""),
                    "reference_only": bool(rec.get("reference_only", False)),
                })
            records.sort(key=lambda r: (r["label"].lower(), r["block"].lower()))

        def refresh(*_):
            tree.delete(*tree.get_children())
            key_by_item.clear()
            q = search_var.get().strip().lower()
            for rec in records:
                hay = (rec["label"] + "\n" + rec["block"]).lower()
                if q and q not in hay:
                    continue
                iid = tree.insert("", "end", values=(rec["label"],))
                key_by_item[iid] = rec

        def show_selected(_event=None):
            sel = tree.selection()
            if not sel:
                return
            rec = key_by_item.get(sel[0])
            if not rec:
                return
            text = rec["block"]
            if rec["source"]:
                text += "\n\nSource: " + rec["source"]
            preview.configure(state="normal")
            preview.delete("1.0", "end")
            preview.insert("1.0", text)
            preview.configure(state="disabled")

        def selected_rec():
            sel = tree.selection()
            return key_by_item.get(sel[0]) if sel else None

        def edit():
            rec = selected_rec()
            if not rec:
                messagebox.showinfo("Select a Ship To", "Select an address first.", parent=win)
                return
            ed = tk.Toplevel(win)
            ed.title("Edit Ship To")
            ed.geometry("650x500")
            ed.transient(win)
            ed.grab_set()
            fr = ttk.Frame(ed, padding=16)
            fr.pack(fill="both", expand=True)
            ttk.Label(fr, text="Verified Ship To", style="Title.TLabel").pack(anchor="w")
            ttk.Label(fr, text="Keep the line breaks exactly as the address should print.",
                      style="Subtitle.TLabel").pack(anchor="w", pady=(2, 10))
            box = scrolledtext.ScrolledText(fr, height=14, wrap="word", font=("Segoe UI", 10))
            box.pack(fill="both", expand=True)
            box.insert("1.0", rec["block"])

            def save_edit():
                val = box.get("1.0", "end").strip()
                lines = [x.strip() for x in val.splitlines() if x.strip()]
                if len(lines) < 2:
                    messagebox.showerror("Complete Ship To",
                                         "Enter the complete multi-line Ship To block.", parent=ed)
                    return
                val = "\n".join(lines)
                # Editing a mapped record updates it. Editing a historical
                # reference saves the corrected block as a new reusable reference.
                self.ship_store.set(
                    rec["key"] if not rec["key"].startswith("library::") else
                    "library::" + __import__("hashlib").sha1(val.encode("utf-8")).hexdigest()[:16],
                    val,
                    source="Edited in Ship To Manager"
                )
                rebuild_records()
                refresh()
                ed.destroy()

            ttk.Button(fr, text="SAVE CHANGES", command=save_edit,
                       style="Primary.TButton").pack(anchor="e", pady=(10, 0))

        def delete():
            rec = selected_rec()
            if not rec:
                messagebox.showinfo("Select a Ship To", "Select an address first.", parent=win)
                return
            if messagebox.askyesno(
                "Remove saved Ship To",
                f"Remove '{rec['label']}' from the bot's saved address list?",
                parent=win
            ):
                self.ship_store.delete(rec["key"])
                rebuild_records()
                refresh()
                preview.configure(state="normal")
                preview.delete("1.0", "end")
                preview.configure(state="disabled")

        tree.bind("<<TreeviewSelect>>", show_selected)
        search_var.trace_add("write", refresh)

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(10, 0))
        ttk.Button(actions, text="Edit Selected", command=edit,
                   style="Secondary.TButton").pack(side="left")
        ttk.Button(actions, text="Remove Selected", command=delete,
                   style="Secondary.TButton").pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Close", command=win.destroy,
                   style="Secondary.TButton").pack(side="right")

        rebuild_records()
        refresh()
        search_entry.focus_set()

    def generate_pdfs(self):
        if not self.result or self.result.hard_errors:
            return

        self.apply_ship_blocks(self.result)
        missing = [
            n for n in self.result.include_nums
            if not self.result.invoices[n].ship_to_block
        ]
        if missing:
            self._show_banner(
                "⚠ One or more invoices still need a verified Ship To.",
                "warn"
            )
            messagebox.showerror(
                "Ship To review required",
                "These invoices still need a verified Ship To:\n" + ", ".join(missing)
            )
            return

        label = period_filename(self.result.period)
        history = load_history()
        prior = [h for h in history if h.get("period") == label]
        if prior:
            last = prior[-1].get("generated_at", "previously")
            if not messagebox.askyesno(
                "PDFs already generated",
                f"Billing PDFs for {label} were already generated ({last}).\n\nGenerate them again?"
            ):
                return

        outdir = filedialog.askdirectory(
            title="Choose where to save the billing PDFs"
        )
        if not outdir:
            return
        outdir = Path(outdir)

        invoice_pdf = outdir / f"{self.cfg.get('invoice_pdf_prefix', 'Open Invoices')} {label}.pdf"
        balance_pdf = outdir / f"{self.cfg.get('balance_pdf_prefix', 'Open Balance')} {label}.pdf"

        try:
            self.generate_btn.configure(state="disabled")
            self._hide_banner()

            self._review_section("PDF GENERATION")

            # Each animated working line becomes the actual final filename.
            self._append_review("Generating invoice PDF...")
            time.sleep(0.65)
            generate_invoice_packet(self.result, self.cfg, invoice_pdf)
            self._replace_last_review_line(f"✓ {invoice_pdf.name}")
            time.sleep(0.30)

            self._append_review("Generating customer open-balance PDF...")
            time.sleep(0.65)
            generate_open_balance(self.result, self.cfg, balance_pdf)
            self._replace_last_review_line(f"✓ {balance_pdf.name}")
            time.sleep(0.30)

            self.generated_paths = [invoice_pdf, balance_pdf]
            self.outlook_draft_opened = False

            history.append({
                "period": label,
                "generated_at": datetime.now().isoformat(timespec="seconds"),
                "invoice_count": len(self.result.include_nums),
                "open_balance": str(self.result.open_total),
                "invoice_pdf": str(invoice_pdf),
                "balance_pdf": str(balance_pdf),
                "tracking_overrides": {
                    n: self.result.invoices[n].tracking_override
                    for n in self.result.include_nums
                    if self.result.invoices[n].tracking_override
                },
            })
            save_history(history)
            technical_log(f"Generated PDFs: {invoice_pdf} | {balance_pdf}")

            self._show_banner(
                "BILLING PDFS CREATED\n"
                "Review both PDFs before moving to the Outlook draft step.",
                "good"
            )
            self._set_button_state("pdfs")

        except Exception as e:
            technical_log("PDF ERROR\n" + traceback.format_exc())
            self._replace_last_review_line("✕ Billing PDF generation failed")
            self._show_banner(
                "PDF GENERATION FAILED\n"
                "Resolve the error and try generating the billing PDFs again.",
                "bad"
            )
            self._set_button_state("ready")
            messagebox.showerror(
                "PDF generation failed",
                f"{e}\n\nTechnical details were saved in the logs folder."
            )

    def _has_missing_tracking(self):
        if not self.result:
            return False
        return any(
            issue.code == "TRACKING_NOT_FOUND"
            for issue in self.result.warnings
        )

    def _outlook_email_body(self):
        lines = [
            "Greetings,",
            "",
            "Please see the attached unpaid invoices at your earliest convenience.",
            "",
        ]

        # Only make the tracking claim when it is actually true.
        if not self._has_missing_tracking():
            lines.extend([
                "All invoices now include a tracking number to help make the process easier on your end.",
                "",
            ])

        lines.extend([
            "If you have any questions or need any adjustments, please let us know.",
        ])
        return "\n".join(lines)

    def _outlook_cc_recipients(self):
        return list(self.cfg.get("email_cc", []))

    def _cc_clipboard_text(self):
        return "; ".join(self._outlook_cc_recipients())

    def _copy_cc_to_clipboard(self):
        """Keep a one-paste CC fallback ready because Outlook Web can ignore cc URLs."""
        try:
            self.clipboard_clear()
            self.clipboard_append(self._cc_clipboard_text())
            self.update()
            return True
        except Exception:
            technical_log("CC CLIPBOARD ERROR\n" + traceback.format_exc())
            return False

    def _outlook_web_url(self):
        """Build a modern Outlook Web compose deeplink."""
        if not self.result:
            return ""

        to_recipients = list(self.cfg.get("email_to", []))
        cc_recipients = self._outlook_cc_recipients()

        subject = f"{self.cfg.get('email_subject_prefix', 'Open Invoices')} {period_filename(self.result.period)}"
        query = urllib.parse.urlencode(
            {
                "to": ";".join(to_recipients),
                "cc": ";".join(cc_recipients),
                "subject": subject,
                "body": self._outlook_email_body(),
            },
            quote_via=urllib.parse.quote
        )
        return f"https://outlook.office.com/mail/deeplink/compose?{query}"

    def _find_edge_executable(self):
        """Return Microsoft Edge executable when available on Windows."""
        if not sys.platform.startswith("win"):
            return None

        candidates = [
            Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Microsoft/Edge/Application/msedge.exe",
            Path(os.environ.get("PROGRAMFILES", "")) / "Microsoft/Edge/Application/msedge.exe",
            Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft/Edge/Application/msedge.exe",
        ]
        for candidate in candidates:
            try:
                if candidate.is_file():
                    return candidate
            except Exception:
                pass
        return None

    def _open_outlook_web_window(self, url):
        """Prefer a separate Edge window; fall back to the default browser."""
        if sys.platform.startswith("win"):
            edge = self._find_edge_executable()
            if edge:
                subprocess.Popen(
                    [str(edge), "--new-window", url],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL
                )
                return True
            os.startfile(url)
            return False

        if sys.platform == "darwin":
            os.system(f'open "{url}"')
        else:
            os.system(f'xdg-open "{url}"')
        return False

    def _tile_outlook_and_folder(self, folder, attempt=0):
        """Best-effort Windows tiling: Outlook Web left, PDF folder right."""
        if not sys.platform.startswith("win"):
            return

        try:
            user32 = ctypes.windll.user32

            rect = wintypes.RECT()
            SPI_GETWORKAREA = 0x0030
            if not user32.SystemParametersInfoW(
                SPI_GETWORKAREA, 0, ctypes.byref(rect), 0
            ):
                return

            windows = []

            @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
            def enum_proc(hwnd, _lparam):
                if not user32.IsWindowVisible(hwnd):
                    return True

                length = user32.GetWindowTextLengthW(hwnd)
                title_buf = ctypes.create_unicode_buffer(max(2, length + 1))
                user32.GetWindowTextW(hwnd, title_buf, len(title_buf))

                class_buf = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(hwnd, class_buf, 256)

                title = title_buf.value.strip()
                class_name = class_buf.value.strip()
                if title:
                    windows.append((hwnd, title, class_name))
                return True

            user32.EnumWindows(enum_proc, 0)

            folder_name = Path(folder).name.lower()

            explorer_candidates = [
                item for item in windows
                if item[2] == "CabinetWClass"
            ]
            explorer = next(
                (
                    item for item in explorer_candidates
                    if folder_name and folder_name in item[1].lower()
                ),
                explorer_candidates[-1] if explorer_candidates else None
            )

            browser_candidates = [
                item for item in windows
                if (
                    "outlook" in item[1].lower()
                    or "microsoft 365" in item[1].lower()
                    or "mail -" in item[1].lower()
                    or "sign in" in item[1].lower()
                )
                and item[2] != "CabinetWClass"
            ]
            browser = browser_candidates[-1] if browser_candidates else None

            if browser is None:
                fg = user32.GetForegroundWindow()
                for item in windows:
                    if item[0] == fg and item[2] != "CabinetWClass":
                        browser = item
                        break

            if explorer is None or browser is None or explorer[0] == browser[0]:
                if attempt < 6:
                    self.after(
                        700,
                        lambda: self._tile_outlook_and_folder(folder, attempt + 1)
                    )
                return

            x = rect.left
            y = rect.top
            total_w = rect.right - rect.left
            total_h = rect.bottom - rect.top

            outlook_w = int(total_w * 0.62)
            folder_w = total_w - outlook_w

            SW_RESTORE = 9
            user32.ShowWindow(browser[0], SW_RESTORE)
            user32.ShowWindow(explorer[0], SW_RESTORE)

            user32.MoveWindow(browser[0], x, y, outlook_w, total_h, True)
            user32.MoveWindow(
                explorer[0],
                x + outlook_w,
                y,
                folder_w,
                total_h,
                True
            )

            try:
                user32.SetForegroundWindow(browser[0])
            except Exception:
                pass

        except Exception:
            technical_log("WINDOW TILING ERROR\n" + traceback.format_exc())

    def _open_generated_folder(self):
        if not self.generated_paths:
            return False

        folder = Path(self.generated_paths[0]).parent
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(folder))
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
            return True
        except Exception:
            technical_log("OUTPUT FOLDER OPEN ERROR\n" + traceback.format_exc())
            return False

    def _show_attachment_reminder(self):
        """Show a non-modal, always-on-top reminder for the one manual email step."""
        if not self.generated_paths:
            return

        if self.attachment_reminder is not None:
            try:
                if self.attachment_reminder.winfo_exists():
                    self.attachment_reminder.lift()
                    self.attachment_reminder.attributes("-topmost", True)
                    return
            except Exception:
                self.attachment_reminder = None

        win = tk.Toplevel(self)
        self.attachment_reminder = win
        win.title("Next Step: Attach Billing PDFs")
        win.configure(bg=COLORS["light"])
        win.resizable(False, False)

        try:
            win.attributes("-topmost", True)
        except Exception:
            pass

        width = 610
        height = 310
        try:
            screen_w = win.winfo_screenwidth()
            x = max(20, (screen_w - width) // 2)
        except Exception:
            x = 120
        y = 85
        win.geometry(f"{width}x{height}+{x}+{y}")

        outer = tk.Frame(
            win,
            bg=COLORS["light"],
            padx=22,
            pady=18
        )
        outer.pack(fill="both", expand=True)

        tk.Label(
            outer,
            text="NEXT STEP: ATTACH BOTH BILLING PDFs",
            bg=COLORS["light"],
            fg=COLORS["navy"],
            font=("Segoe UI", 13, "bold"),
            anchor="w"
        ).pack(fill="x")

        tk.Label(
            outer,
            text=(
                "Drag these 2 files from the folder that just opened into the Outlook Web email:"
            ),
            bg=COLORS["light"],
            fg=COLORS["text"],
            font=("Segoe UI", 10),
            justify="left",
            anchor="w",
            wraplength=560
        ).pack(fill="x", pady=(10, 10))

        files_box = tk.Frame(
            outer,
            bg="#FFFFFF",
            highlightbackground=COLORS["border"],
            highlightthickness=1,
            padx=12,
            pady=10
        )
        files_box.pack(fill="x")

        for p in self.generated_paths:
            tk.Label(
                files_box,
                text=f"✓ {Path(p).name}",
                bg="#FFFFFF",
                fg=COLORS["navy"],
                font=("Segoe UI", 9, "bold"),
                justify="left",
                anchor="w",
                wraplength=540
            ).pack(fill="x", pady=2)

        tk.Label(
            outer,
            text="Then review the email and send.",
            bg=COLORS["light"],
            fg=COLORS["text"],
            font=("Segoe UI", 10, "bold"),
            anchor="w"
        ).pack(fill="x", pady=(10, 4))

        tk.Label(
            outer,
            text=(
                f"CC: {self.cfg.get('email_cc_display', 'see config.json')}. "
                "If Outlook leaves CC blank, click Cc and press Ctrl+V — the CC list is already copied."
            ),
            bg=COLORS["light"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8),
            anchor="w",
            justify="left",
            wraplength=560
        ).pack(fill="x", pady=(0, 10))

        buttons = ttk.Frame(outer)
        buttons.pack(fill="x")
        buttons.columnconfigure(0, weight=1)
        buttons.columnconfigure(1, weight=1)

        def got_it():
            try:
                win.destroy()
            finally:
                self.attachment_reminder = None

        ttk.Button(
            buttons,
            text="OPEN PDF FOLDER",
            command=self._open_generated_folder,
            style="Secondary.TButton"
        ).grid(row=0, column=0, sticky="w")

        ttk.Button(
            buttons,
            text="GOT IT",
            command=got_it,
            style="Primary.TButton"
        ).grid(row=0, column=1, sticky="e")

        win.protocol("WM_DELETE_WINDOW", got_it)

    def create_outlook_draft(self):
        """Open a prefilled email directly in Outlook on the web.

        This avoids the desktop New Outlook licensing problem on machines where
        the Windows app is unavailable. It still requires no Graph registration,
        API token, or bot-side Microsoft 365 authentication. Attachments remain
        a deliberate human step.
        """
        if not self.result or not self.generated_paths:
            messagebox.showerror(
                "Billing PDFs required",
                "Generate and review both billing PDFs before creating the Outlook draft."
            )
            return

        missing_files = [
            str(p) for p in self.generated_paths
            if not Path(p).exists()
        ]
        if missing_files:
            self._show_banner(
                "OUTLOOK DRAFT NOT OPENED\n"
                "One or both billing PDFs could not be found. Generate the PDFs again.",
                "bad"
            )
            messagebox.showerror(
                "Billing PDFs not found",
                "One or both billing PDFs could not be found:\n\n"
                + "\n".join(missing_files)
            )
            return

        if self.outlook_draft_opened:
            if not messagebox.askyesno(
                "Open another Outlook draft?",
                "An Outlook draft was already opened for this billing run.\n\n"
                "Open another one?"
            ):
                return

        # Open the PDF folder first so it is ready behind the compose window.
        folder_opened = self._open_generated_folder()

        # Keep the expected CC list on the clipboard as a backup. Outlook Web's
        # URL compose flow is inconsistent with CC on some Microsoft 365 tenants.
        self._copy_cc_to_clipboard()

        try:
            outlook_web_url = self._outlook_web_url()
            self._open_outlook_web_window(outlook_web_url)
        except Exception as e:
            technical_log("OUTLOOK WEB HANDOFF ERROR\n" + traceback.format_exc())
            self._show_banner(
                "OUTLOOK DRAFT NOT OPENED\n"
                "The browser could not open Outlook on the web.",
                "bad"
            )
            messagebox.showerror(
                "Could not open Outlook on the web",
                "The bot could not open Outlook on the web in your browser.\n\n"
                "Make sure you have a default web browser, then try again.\n\n"
                f"Details: {e}"
            )
            return

        already_opened = self.outlook_draft_opened
        self.outlook_draft_opened = True

        if not already_opened:
            self._review_section("OUTLOOK DRAFT")
            self._append_review("✓ Outlook Web draft opened")
        else:
            self._replace_progress_line_starting(
                "✓ Outlook Web draft opened",
                "✓ Outlook Web draft opened again"
            )

        tracking_note = (
            " The tracking sentence was omitted because one or more tracking numbers are still missing."
            if self._has_missing_tracking()
            else ""
        )
        self._show_banner(
            "OUTLOOK WEB DRAFT OPENED\n"
            "Next: Attach both billing PDFs from the open folder, then review and send."
            + tracking_note,
            "good"
        )
        self._set_button_state("outlook")

        # Prefer a dedicated Outlook Web window and place it beside the PDF
        # folder for easy drag-and-drop.
        if sys.platform.startswith("win") and self.generated_paths:
            folder = Path(self.generated_paths[0]).parent
            self.after(
                1700,
                lambda: self._tile_outlook_and_folder(folder)
            )

        # Show the non-modal attachment reminder after both windows open.
        self.after(2500, self._show_attachment_reminder)

    def open_pdfs(self):
        if not self.generated_paths:
            return

        opened_any = False
        for p in self.generated_paths:
            try:
                if sys.platform.startswith("win"):
                    os.startfile(str(p))
                elif sys.platform == "darwin":
                    os.system(f'open "{p}"')
                else:
                    os.system(f'xdg-open "{p}"')
                opened_any = True
            except Exception:
                pass

        if opened_any:
            self._show_banner(
                "PDFS OPENED FOR REVIEW\n"
                "Confirm both PDFs look correct before creating the Outlook draft.",
                "good"
            )
            self._set_button_state("reviewed")


def main():
    BillingApp().mainloop()


if __name__ == "__main__":
    main()
