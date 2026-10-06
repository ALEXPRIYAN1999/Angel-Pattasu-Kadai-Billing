"""
ANGEL PATTASU KADAI - Billing Desktop App (FINAL)
- Billing: search -> Tab to qty -> add, fully editable cart
- Cart columns: S.No | Product | Qty | Orig Rate | Disc % | Rate | Amount
- Double-click cart row to edit EVERYTHING except Amount (auto-calc)
- PDF: Original Rate, Discount %, Final Rate, Amount + 80% discount footer
Run: python angel_billing.py
Install: pip install reportlab pillow
"""

import os, json, re, csv, datetime, copy
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None

# =========================================================
# PATHS
# =========================================================
BASE_DIR      = os.path.dirname(os.path.abspath(__file__))
# Store application data inside the current Windows user's home folder.
# Example: C:\\Users\\LENOVO\\angel_data
# This automatically becomes C:\\Users\\ACER\\angel_data, etc. on another PC.
USER_HOME     = os.path.expanduser("~")
DATA_DIR      = os.path.join(USER_HOME, "angel_data")
PDF_DIR       = os.path.join(DATA_DIR, "invoices")
# Keep Statement PDFs separate from invoice PDFs. JSON/database files are unchanged.
STATEMENT_DIR = os.path.join(DATA_DIR, "Statement")
# Keep each document type in its own folder so the same numeric sequence
# can safely exist across Tax Invoice / Bill Of Supply / Estimate.
BILL_TYPE_CONFIG = {
    "Tax Invoice": {"prefix": "TIANG", "folder": "Tax_Invoice"},
    "Bill Of Supply": {"prefix": "BOAANG", "folder": "Bill_Of_Supply"},
    "Estimate": {"prefix": "ESTANG", "folder": "Estimate"},
}
BILL_TYPE_PDF_DIRS = {
    bill_type: os.path.join(DATA_DIR, cfg["folder"])
    for bill_type, cfg in BILL_TYPE_CONFIG.items()
}
for _bill_dir in BILL_TYPE_PDF_DIRS.values():
    os.makedirs(_bill_dir, exist_ok=True)
ASSET_DIR     = os.path.join(BASE_DIR, "assets")
LOGO_PATH     = os.path.join(ASSET_DIR, "angel_logo.png")
LOGO_FALLBACK_PATH = os.path.join(ASSET_DIR, "angel_logo.jpeg")
ROUND_LOGO_PATH = os.path.join(ASSET_DIR, "logo_round.png")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(PDF_DIR, exist_ok=True)
os.makedirs(STATEMENT_DIR, exist_ok=True)
os.makedirs(ASSET_DIR, exist_ok=True)

PRODUCTS_FILE = os.path.join(DATA_DIR, "products.json")
BILLS_FILE    = os.path.join(DATA_DIR, "bills.json")
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
STOCK_FILE    = os.path.join(DATA_DIR, "stock.json")

# =========================================================
# SETTINGS
# =========================================================
DEFAULT_SETTINGS = {
    "shop_name":  "ANGEL PATTASU KADAI",
    "address":    "AYYANAR NAGAR, NH-7, PATTAMPUDHUR",
    "phone":      "+91 82208 02867",
    "email":      "angelpattasukadai08@gmail.com",
    "gstin":      "33ABRFA4846J1Z3",
    "state":      "Tamil Nadu (33)",
    "tagline":    "Most Reliable Brand",
    "next_bill":  1,
    "bill_prefix": "ANG",
    "rate_edit_password": "admin",
    "default_discount_pct": 80,
    "theme": "Dark Blue",
}

# =========================================================
# STORAGE
# =========================================================
def load_json(path, default):
    if not os.path.exists(path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(default, f, indent=2, ensure_ascii=False)
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default

def save_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    os.replace(tmp, path)

SETTINGS = load_json(SETTINGS_FILE, DEFAULT_SETTINGS)
PRODUCTS = load_json(PRODUCTS_FILE, [])
BILLS    = load_json(BILLS_FILE, [])
STOCK     = load_json(STOCK_FILE, [])

# =========================================================
# NUMBER → WORDS
# =========================================================
def num_to_words(n):
    try: n = int(round(n))
    except: return ""
    if n == 0: return "Zero Rupees Only"
    ones = ["","One","Two","Three","Four","Five","Six","Seven","Eight","Nine",
            "Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen",
            "Seventeen","Eighteen","Nineteen"]
    tens = ["","","Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"]
    def two(n):
        if n < 20: return ones[n]
        return tens[n//10] + (" " + ones[n%10] if n%10 else "")
    def three(n):
        if n < 100: return two(n)
        return ones[n//100] + " Hundred" + (" " + two(n%100) if n%100 else "")
    parts = []
    crore, n = divmod(n, 10000000)
    lakh,  n = divmod(n, 100000)
    thou,  n = divmod(n, 1000)
    hund,  n = divmod(n, 100)
    if crore: parts.append(three(crore)+" Crore")
    if lakh:  parts.append(three(lakh)+" Lakh")
    if thou:  parts.append(three(thou)+" Thousand")
    if hund:  parts.append(two(hund))
    return " ".join(parts).strip() + " Rupees Only"

# =========================================================
# A4 PDF INVOICE
# =========================================================
def _safe_number(value, default=0.0):
    """Read a number from normal text, including values such as 9% or ₹1,000."""
    try:
        if isinstance(value, (int, float)):
            return float(value)
        s = str(value or "").strip().replace(",", "").replace("₹", "").replace("Rs.", "")
        import re as _re
        m = _re.search(r"-?\d+(?:\.\d+)?", s)
        return float(m.group(0)) if m else float(default)
    except (TypeError, ValueError):
        return float(default)


def _bill_totals(bill):
    """Return invoice totals. User 1 keeps the original calculation; User 2 adds GST.

    User 2 normally calculates Total GST from CGST + SGST.  If the user has
    explicitly edited Total GST in the billing screen, that entered amount is
    used for the invoice total and PDF instead.
    """
    total_orig = 0.0
    total_discounted = 0.0
    for it in bill.get("items", []) or []:
        qty = _safe_number(it.get("qty", 0))
        orig = _safe_number(it.get("orig_rate", it.get("rate", 0)))
        rate = _safe_number(it.get("rate", 0))
        total_orig += qty * orig
        total_discounted += qty * rate

    total_discount = total_orig - total_discounted
    packing = max(0.0, _safe_number(bill.get("packing_charge", 0)))
    taxable = total_discounted + packing

    user = str(bill.get("billing_user", "User 1") or "User 1")
    if user == "User 2":
        cgst_pct = max(0.0, min(100.0, _safe_number(bill.get("cgst_pct", 9), 9)))
        sgst_pct = max(0.0, min(100.0, _safe_number(bill.get("sgst_pct", 9), 9)))
    else:
        cgst_pct = sgst_pct = 0.0

    cgst = round(taxable * cgst_pct / 100.0, 2)
    sgst = round(taxable * sgst_pct / 100.0, 2)
    calculated_gst = round(cgst + sgst, 2)

    # Total GST is normally automatic.  User 2 may override it manually.
    manual_gst = bool(bill.get("total_gst_manual", False)) and user == "User 2"
    if manual_gst:
        total_gst = max(0.0, round(_safe_number(
            bill.get("total_gst_override", calculated_gst),
            calculated_gst
        ), 2))
    else:
        total_gst = calculated_gst

    before_round = taxable + total_gst
    grand = round(before_round)
    roundoff = grand - before_round

    return {
        "billing_user": user,
        "original_value": round(total_orig, 2),
        "discount_value": round(total_discount, 2),
        "sub_total": round(total_discounted, 2),
        "packing_charge": round(packing, 2),
        "taxable_value": round(taxable, 2),
        "cgst_pct": cgst_pct,
        "sgst_pct": sgst_pct,
        "cgst": cgst,
        "sgst": sgst,
        "calculated_gst": calculated_gst,
        "total_gst_manual": manual_gst,
        "total_gst": total_gst,
        "before_round": round(before_round, 2),
        "roundoff": round(roundoff, 2),
        "grand_total": float(grand),
    }


def generate_user3_transport_invoice(bill, out_path):
    """Generate User 3's plain transport-purpose Bill Of Supply PDF.

    User 3 is a credit document: no Paid/Balance/Not Paid/Partial fields are
    printed. The document shows only CREDIT and a signature area at the end.
    """
    c = canvas.Canvas(out_path, pagesize=A4)
    W, H = A4
    totals = _bill_totals(bill)

    # Plain black-and-white document: no shop logo, address, GST, watermark,
    # coloured header or payment/balance details.
    c.setFillColor(colors.black)
    c.setStrokeColor(colors.black)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W/2, H-18*mm, "BILL OF SUPPLY")

    c.setFont("Helvetica-Bold", 8.5)
    y = H - 29*mm
    c.drawString(18*mm, y, f"Bill No : {bill.get('bill_no', '-')}")
    c.drawString(18*mm, y-5*mm, f"Date : {bill.get('date', '-')}")
    c.drawString(18*mm, y-10*mm, f"Customer : {bill.get('customer') or '-'}")
    c.drawString(18*mm, y-15*mm, f"Mobile : {bill.get('mobile') or '-'}")

    right_x = 112*mm
    c.drawString(right_x, y, "Address :")
    c.setFont("Helvetica", 8)
    addr_lines = str(bill.get('address') or '-').replace('\\r', '').split('\\n')[:3]
    for i, ln in enumerate(addr_lines):
        c.drawString(right_x, y-4.5*mm-i*4*mm, str(ln)[:52])

    # Explicit transport credit status.
    c.setFont("Helvetica-Bold", 10)
    c.drawString(right_x, y-18*mm, "CREDIT")

    table_left = 15*mm
    table_right = W-15*mm
    col_w = [9*mm, 65*mm, 13*mm, 19*mm, 15*mm, 24*mm, 35*mm]
    col_labels = ["No.", "Product", "Qty", "Price", "Disc%", "Dis Rate", "Amount"]
    header_h = 6*mm
    y_table = y - 24*mm
    header_bottom = y_table - header_h
    c.setLineWidth(0.7)
    c.rect(table_left, header_bottom, table_right-table_left, header_h, fill=0, stroke=1)
    c.setFont("Helvetica-Bold", 7.2)
    x = table_left
    for w, label in zip(col_w, col_labels):
        if label == "Product":
            c.drawString(x+2*mm, header_bottom+1.9*mm, label)
        else:
            c.drawCentredString(x+w/2, header_bottom+1.9*mm, label)
        c.line(x, header_bottom, x, header_bottom+header_h)
        x += w
    c.line(table_right, header_bottom, table_right, header_bottom+header_h)

    body_font = "Helvetica"
    body_size = 7.0
    line_h = 3.15*mm
    min_row_h = 4.15*mm
    product_inner_w = col_w[1] - 4*mm

    def wrap_text(text, max_width, max_lines=2):
        text = str(text or "-").replace("\\n", " ").strip() or "-"
        words = text.split()
        lines, cur = [], ""
        for word in words:
            candidate = word if not cur else cur + " " + word
            if stringWidth(candidate, body_font, body_size) <= max_width:
                cur = candidate
                continue
            if cur:
                lines.append(cur); cur = word
            else:
                piece = ""
                for ch in word:
                    test = piece + ch
                    if stringWidth(test, body_font, body_size) <= max_width:
                        piece = test
                    else:
                        if piece: lines.append(piece)
                        piece = ch
                cur = piece
        if cur: lines.append(cur)
        if len(lines) <= max_lines: return lines
        kept = lines[:max_lines]
        last = kept[-1]
        while last and stringWidth(last + "...", body_font, body_size) > max_width:
            last = last[:-1]
        kept[-1] = (last.rstrip() + "...") if last else "..."
        return kept

    def draw_header(top_y, continuation=False):
        if continuation:
            c.setFont("Helvetica-Bold", 10)
            c.drawCentredString(W/2, H-12*mm, "BILL OF SUPPLY - CONTINUED")
            c.setFont("Helvetica", 7.5)
            c.drawCentredString(W/2, H-16.5*mm,
                                f"Bill No: {bill.get('bill_no', '-')}   Date: {bill.get('date', '-')}")
            top_y = H-20.5*mm
        bottom = top_y-header_h
        c.rect(table_left, bottom, table_right-table_left, header_h, fill=0, stroke=1)
        c.setFont("Helvetica-Bold", 7.2)
        xx = table_left
        for ww, label in zip(col_w, col_labels):
            if label == "Product":
                c.drawString(xx+2*mm, bottom+1.9*mm, label)
            else:
                c.drawCentredString(xx+ww/2, bottom+1.9*mm, label)
            c.line(xx, bottom, xx, bottom+header_h)
            xx += ww
        c.line(table_right, bottom, table_right, bottom+header_h)
        c.setFont(body_font, body_size)
        return bottom

    ry = header_bottom
    page_items = 0
    items = bill.get("items", []) or []
    page_break_y = 72*mm
    for i, it in enumerate(items, 1):
        orig = _safe_number(it.get("orig_rate", it.get("rate", 0)))
        pct = _safe_number(it.get("disc_pct", 0))
        qty = _safe_number(it.get("qty", 0))
        rate = _safe_number(it.get("rate", 0))
        amt = qty * rate
        product_lines = wrap_text(it.get("name", ""), product_inner_w)
        row_h = max(min_row_h, len(product_lines)*line_h + 1.0*mm)
        if page_items and ry-row_h < page_break_y:
            c.line(table_left, ry, table_right, ry)
            c.showPage()
            c.setFillColor(colors.black); c.setStrokeColor(colors.black)
            ry = draw_header(H-20.5*mm, continuation=True)
            page_items = 0
        c.line(table_left, ry-row_h, table_right, ry-row_h)
        # vertical cell borders
        xx = table_left
        for ww in col_w:
            c.line(xx, ry, xx, ry-row_h); xx += ww
        c.line(table_right, ry, table_right, ry-row_h)
        baseline = ry-row_h/2-0.8*mm if len(product_lines)==1 else ry-2.45*mm
        x = table_left
        vals = [str(i), None, f"{qty:g}", f"{orig:.2f}", f"{pct:g}%", f"{rate:.2f}", f"{amt:.2f}"]
        for w, label, val in zip(col_w, col_labels, vals):
            if label == "Product":
                for li, line in enumerate(product_lines):
                    c.drawString(x+2*mm, baseline-li*line_h, line)
            else:
                c.drawCentredString(x+w/2, baseline, val)
            x += w
        ry -= row_h
        page_items += 1
    c.line(table_left, ry, table_right, ry)

    ty = ry - 5*mm
    label_right = W-58*mm
    value_x = W-15*mm
    c.setFont("Helvetica", 8)
    rows = [
        ("Sub Total Value", totals["original_value"]),
        (f"Total Discount @ {int(round((totals['discount_value']/totals['original_value']*100) if totals['original_value'] else 0))}%", totals["discount_value"]),
        ("Sub Total", totals["sub_total"]),
        ("Packing Charges", totals["packing_charge"]),
        ("Round Off", totals["roundoff"]),
    ]
    for label, value in rows:
        c.drawRightString(label_right, ty, label)
        c.drawRightString(value_x, ty, f"Rs. {value:,.2f}")
        ty -= 4.25*mm
    c.setFont("Helvetica-Bold", 9)
    c.drawRightString(label_right, ty, "GRAND TOTAL")
    c.drawRightString(value_x, ty, f"Rs. {totals['grand_total']:,.2f}")
    ty -= 6*mm
    c.setFont("Helvetica-Oblique", 7.3)
    c.drawString(18*mm, ty, f"Amount in words: {num_to_words(totals['grand_total'])}")

    # Credit only: deliberately no Paid / Balance / Payment amount fields.
    c.setFont("Helvetica-Bold", 9)
    c.drawString(18*mm, 35*mm, "CREDIT")

    # User 3: no signature. Keep the transport credit document clean.
    c.showPage()
    c.save()
    return out_path


def generate_a4_invoice(bill, out_path):
    # User 3 is a transport-purpose credit Bill Of Supply. Keep all existing
    # User 1/User 2 PDF logic unchanged and route only User 3 here.
    if str(bill.get("billing_user", "")).strip() == "User 3":
        return generate_user3_transport_invoice(bill, out_path)

    """Generate a compact, paper-saving A4 billing document.

    Layout goals:
      * Compact header/info block so the item table starts higher.
      * Target 30 normal product rows on one A4 page.
      * Long product names wrap only inside the Product column.
      * Never allow text to cross into neighbouring columns.
      * Keep totals, Grand Total and footer separated/aligned.
      * Continue automatically to another A4 page for genuinely large bills.
    """
    c = canvas.Canvas(out_path, pagesize=A4)
    W, H = A4
    totals = _bill_totals(bill)

    # ---- Brand logos: use the supplied PNG first, JPEG as a safe fallback. ----
    logo_path = LOGO_PATH if os.path.exists(LOGO_PATH) else LOGO_FALLBACK_PATH
    if os.path.exists(logo_path):
        try:
            c.drawImage(logo_path, 15*mm, H-34*mm,
                        width=35*mm, height=20*mm,
                        preserveAspectRatio=True, anchor='sw', mask="auto")
        except Exception:
            pass

    if os.path.exists(ROUND_LOGO_PATH):
        try:
            c.drawImage(ROUND_LOGO_PATH, W-42*mm, H-35*mm,
                        width=26*mm, height=26*mm,
                        preserveAspectRatio=True, anchor='sw', mask="auto")
        except Exception:
            pass

    # ---- Compact header ----
    c.setFillColor(colors.HexColor("#8B0000"))
    c.setFont("Helvetica-Bold", 19)
    c.drawCentredString(W/2, H-14*mm, SETTINGS["shop_name"])
    c.setFillColor(colors.HexColor("#B8860B"))
    c.setFont("Helvetica-Oblique", 9)
    c.drawCentredString(W/2, H-19*mm, SETTINGS["tagline"])
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 8.2)
    c.drawCentredString(W/2, H-24*mm, SETTINGS["address"])
    c.drawCentredString(W/2, H-28.5*mm,
                        f"Phone: {SETTINGS['phone']}   |   Email: {SETTINGS['email']}")
    if str(bill.get("billing_user", "")).strip() != "User 4":
        c.drawCentredString(W/2, H-33*mm,
                            f"GSTIN: {SETTINGS['gstin']}   |   State: {SETTINGS['state']}")

    c.setStrokeColor(colors.HexColor("#8B0000"))
    c.setLineWidth(1.0)
    c.line(15*mm, H-36.5*mm, W-15*mm, H-36.5*mm)

    # ---- Document type title ----
    # The billing screen lets the user choose Bill Of Supply, Tax Invoice,
    # or Estimate.  Keep the selected type in the saved bill so reprints
    # and history PDFs use the same title.
    document_type = str(bill.get("bill_type", "Bill Of Supply") or "Bill Of Supply").strip()
    if document_type not in ("Bill Of Supply", "Tax Invoice", "Estimate"):
        document_type = "Bill Of Supply"

    c.setFillColor(colors.HexColor("#8B0000"))
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(W/2, H-43*mm, document_type)

    # ---- Compact bill information ----
    y = H - 51*mm
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 8.5)
    c.drawString(20*mm, y,       f"Bill No : {bill.get('bill_no', '-')}")
    c.drawString(20*mm, y-5*mm,  f"Date : {bill.get('date', '-')}")
    c.drawString(20*mm, y-10*mm, f"Customer : {bill.get('customer') or '-'}")
    c.drawString(20*mm, y-15*mm, f"Mobile : {bill.get('mobile') or '-'}")

    # Address is compact and bounded to the right half of the page.
    right_x = 112*mm
    c.drawString(right_x, y, "Address :")
    c.setFont("Helvetica", 8)
    addr_lines = str(bill.get('address') or '-').replace('\r', '').split('\n')[:2]
    for i, ln in enumerate(addr_lines):
        c.drawString(right_x, y-4.5*mm-i*4*mm, str(ln)[:52])
    # Customer GSTIN is intentionally not displayed below the address.
    # The shop GSTIN in the header remains unchanged.
    total_due = float(totals['grand_total'])
    paid_amt = max(0.0, min(total_due, _safe_number(bill.get('amount_paid', 0))))
    if total_due > 0 and paid_amt >= total_due - 0.005:
        paid_amt = total_due
        pay_status = "Paid"
    elif paid_amt <= 0.005:
        paid_amt = 0.0
        pay_status = "Not Paid"
    else:
        pay_status = "Partial"
    balance_amt = max(0.0, total_due - paid_amt)
    if str(bill.get("billing_user", "")).strip() not in ("User 3", "User 4"):
        c.drawString(right_x, y-18*mm,
                     f"Payment : {pay_status}   Paid: Rs. {paid_amt:,.2f}   Balance: Rs. {balance_amt:,.2f}")
    elif str(bill.get("billing_user", "")).strip() == "User 4":
        c.setFont("Helvetica-Bold", 10)
        c.drawString(right_x, y-18*mm, "CREDIT")

    # ---- Compact table ----
    body_font = "Helvetica"
    body_size = 7.0
    table_left = 15*mm
    table_right = W-15*mm
    # 180 mm total width. Product gets the largest safe text area.
    col_w = [9*mm, 65*mm, 13*mm, 19*mm, 15*mm, 24*mm, 35*mm]
    col_labels = ["No.", "Product", "Qty", "Price", "Disc%", "Dis Rate", "Amount"]
    y_table = y - 23*mm

    header_h = 6*mm
    header_top = y_table
    header_bottom = header_top - header_h
    c.setFillColor(colors.HexColor("#8B0000"))
    c.rect(table_left, header_bottom, table_right-table_left, header_h, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 7.2)
    x = table_left
    for w, label in zip(col_w, col_labels):
        if label == "Product":
            c.drawString(x+2*mm, header_bottom+1.9*mm, label)
        else:
            c.drawCentredString(x+w/2, header_bottom+1.9*mm, label)
        x += w

    # Faint centre watermark; it is behind the table text.
    if os.path.exists(ROUND_LOGO_PATH):
        try:
            c.saveState()
            if hasattr(c, "setFillAlpha"):
                c.setFillAlpha(0.055)
            c.drawImage(ROUND_LOGO_PATH, W/2-30*mm, H/2-30*mm,
                        width=60*mm, height=60*mm,
                        preserveAspectRatio=True, anchor='sw', mask="auto")
            c.restoreState()
        except Exception:
            pass
    c.setFillColor(colors.black)
    c.setFont(body_font, body_size)

    def wrap_text(text, font_name, font_size, max_width, max_lines=2):
        """Wrap text by measured PDF width, including unbroken product names."""
        text = str(text or "-").replace("\n", " ").strip()
        if not text:
            return ["-"]
        words = text.split()
        lines = []
        cur = ""
        for word in words:
            candidate = word if not cur else cur + " " + word
            if stringWidth(candidate, font_name, font_size) <= max_width:
                cur = candidate
                continue
            if cur:
                lines.append(cur)
                cur = word
            else:
                piece = ""
                for ch in word:
                    test = piece + ch
                    if stringWidth(test, font_name, font_size) <= max_width:
                        piece = test
                    else:
                        if piece:
                            lines.append(piece)
                        piece = ch
                cur = piece
        if cur:
            lines.append(cur)
        if len(lines) <= max_lines:
            return lines
        kept = lines[:max_lines]
        last = kept[-1]
        while last and stringWidth(last + "...", font_name, font_size) > max_width:
            last = last[:-1]
        kept[-1] = (last.rstrip() + "...") if last else "..."
        return kept

    items = bill.get("items", []) or []
    line_h = 3.15*mm
    min_row_h = 4.15*mm
    product_inner_w = col_w[1] - 4*mm
    # Totals/footers need roughly 47 mm. A 30-row normal table therefore has
    # enough room because 30 * 4.15 mm = 124.5 mm.
    page_break_y = 60*mm

    def row_data(item):
        orig = float(item.get("orig_rate", item.get("rate", 0)) or 0)
        pct = float(item.get("disc_pct", 0) or 0)
        qty = float(item.get("qty", 0) or 0)
        rate = float(item.get("rate", 0) or 0)
        amt = qty * rate
        product_lines = wrap_text(item.get("name", ""), body_font, body_size,
                                  product_inner_w, max_lines=2)
        row_h = max(min_row_h, len(product_lines)*line_h + 1.0*mm)
        return row_h, orig, pct, qty, rate, amt, product_lines

    def draw_continuation_header():
        c.setFillColor(colors.HexColor("#8B0000"))
        c.setFont("Helvetica-Bold", 10)
        c.drawCentredString(W/2, H-12*mm,
                            f"{document_type.upper()} - CONTINUED")
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 7.5)
        c.drawCentredString(W/2, H-16.5*mm,
                            f"Bill No: {bill.get('bill_no', '-')}   Date: {bill.get('date', '-')}")
        top = H-20.5*mm
        bottom = top-header_h
        c.setFillColor(colors.HexColor("#8B0000"))
        c.rect(table_left, bottom, table_right-table_left, header_h, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 7.2)
        xx = table_left
        for ww, label in zip(col_w, col_labels):
            if label == "Product":
                c.drawString(xx+2*mm, bottom+1.9*mm, label)
            else:
                c.drawCentredString(xx+ww/2, bottom+1.9*mm, label)
            xx += ww
        c.setFillColor(colors.black)
        c.setFont(body_font, body_size)
        return bottom

    ry = header_bottom
    page_items = 0
    for i, it in enumerate(items, 1):
        row_h, orig, pct, qty, rate, amt, product_lines = row_data(it)
        if page_items and ry - row_h < page_break_y:
            c.setStrokeColor(colors.black); c.setLineWidth(0.5)
            c.line(table_left, ry, table_right, ry)
            c.showPage()
            ry = draw_continuation_header()
            page_items = 0

        c.setStrokeColor(colors.HexColor("#dddddd")); c.setLineWidth(0.2)
        c.line(table_left, ry-row_h, table_right, ry-row_h)

        if len(product_lines) == 1:
            baseline = ry - row_h/2 - 0.8*mm
        else:
            baseline = ry - 2.45*mm

        x = table_left
        vals = [str(i), None, f"{qty:g}", f"{orig:.2f}", f"{pct:g}%", f"{rate:.2f}", f"{amt:.2f}"]
        for w, label, val in zip(col_w, col_labels, vals):
            if label == "Product":
                for li, line in enumerate(product_lines):
                    c.drawString(x+2*mm, baseline-li*line_h, line)
            else:
                c.drawCentredString(x+w/2, baseline, val)
            x += w
        ry -= row_h
        page_items += 1

    c.setStrokeColor(colors.black); c.setLineWidth(0.5)
    c.line(table_left, ry, table_right, ry)

    # ---- Totals: compact and aligned ----
    ty = ry - 4.0*mm
    label_right = W - 58*mm
    value_x = W - 15*mm
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 8.0)
    rows = [
        ("Sub Total Value", totals["original_value"]),
        (f"Total Discount @ {int(round((totals['discount_value']/totals['original_value']*100) if totals['original_value'] else 0))}%", totals["discount_value"]),
        ("Sub Total", totals["sub_total"]),
        ("Packing Charges", totals["packing_charge"]),
    ]
    if totals.get("billing_user") == "User 2":
        rows.extend([
            (f"CGST @ {totals['cgst_pct']:g}%", totals["cgst"]),
            (f"SGST @ {totals['sgst_pct']:g}%", totals["sgst"]),
            
        ])
    rows.append(("Round Off", totals["roundoff"]))
    for label, value in rows:
        c.drawRightString(label_right, ty, label)
        c.drawRightString(value_x, ty, f"Rs. {value:,.2f}")
        ty -= 4.25*mm

    grand_y = ty - 0.5*mm
    label_x = 112*mm
    bar_h = 6.8*mm
    c.setFillColor(colors.HexColor("#8B0000"))
    c.rect(label_x, grand_y-bar_h+0.8*mm, value_x-label_x, bar_h, fill=1, stroke=0)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 9.2)
    c.drawString(label_x+2*mm, grand_y-2.65*mm, "GRAND TOTAL")
    c.drawRightString(value_x-2*mm, grand_y-2.65*mm,
                      f"Rs. {totals['grand_total']:,.2f}")

    c.setFillColor(colors.HexColor("#B8860B"))
    c.setFont("Helvetica-Bold", 7.4)
    c.drawString(20*mm, grand_y-9.0*mm,
                 f"Original Value: Rs. {totals['original_value']:,.2f}   |   "
                 f"Total Discount: Rs. {totals['discount_value']:,.2f}   |   "
                 f"You Saved: Rs. {totals['discount_value']:,.2f}")
    if totals.get("billing_user") == "User 2":
        c.drawString(20*mm, grand_y-13.0*mm,
                     f"CGST {totals['cgst_pct']:g}%: Rs. {totals['cgst']:,.2f}   |   "
                     f"SGST {totals['sgst_pct']:g}%: Rs. {totals['sgst']:,.2f}   |   "
                     f"Total GST: Rs. {totals['total_gst']:,.2f}")

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Oblique", 7.3)
    c.drawString(20*mm, grand_y-18*mm,
                 f"Amount in words: {num_to_words(totals['grand_total'])}")

    # User 1 only: required composition-taxable-person note at the bottom of the billing document.
    if totals.get("billing_user", "User 1") == "User 1":
        c.setFillColor(colors.HexColor("#8B0000"))
        c.setFont("Helvetica-Bold", 7.2)
        c.drawString(20*mm, grand_y-22.5*mm, "Note :")
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 7.0)
        c.drawString(32*mm, grand_y-22.5*mm,
                      "1. I am a Composition Taxable Person, not eligible to collect tax on supplies.")

    c.setStrokeColor(colors.HexColor("#8B0000")); c.setLineWidth(0.6)
    c.line(15*mm, 19*mm, W-15*mm, 19*mm)
    c.setFillColor(colors.HexColor("#8B0000")); c.setFont("Helvetica-Bold", 8.5)
    c.drawCentredString(W/2, 13.5*mm, "Thank You! Visit Again")
    c.setFillColor(colors.HexColor("#666666")); c.setFont("Helvetica", 6.5)
    c.drawCentredString(W/2, 9*mm,
                        f"{SETTINGS['shop_name']}  |  {SETTINGS['phone']}  |  {SETTINGS['email']}")

    # Signature is placed clearly ABOVE the footer separator so it never
    # touches or overlaps the footer line. Applies to User 1, User 2 and User 4.
    c.setStrokeColor(colors.black); c.setLineWidth(0.7)
    sig_x = W-65*mm
    c.line(sig_x, 30*mm, W-18*mm, 30*mm)
    c.setFillColor(colors.black); c.setFont("Helvetica", 7.5)
    c.drawCentredString((sig_x + W-18*mm)/2, 26*mm, "Authorized Signature")
    c.showPage()
    c.save()
    return out_path


# =========================================================
# CSV IMPORT
# =========================================================
def import_from_csv(csv_path):
    products = []

    def parse_money(text):
        if text is None: return None
        s = str(text).strip()
        if not s: return None
        s = s.replace("₹", "").replace(",", "").replace("Rs.", "").strip()
        m = re.search(r"\d+(?:\.\d+)?", s)
        if not m: return None
        try: return float(m.group(0))
        except ValueError: return None

    def clean_unit(text):
        if not text: return "Box"
        m = re.search(r"\b(Box|Pkt\.?|Pkts?|Pcs?|Nos?|Packet)\b", str(text), re.IGNORECASE)
        if not m: return "Box"
        u = m.group(1).rstrip(".").strip().lower()
        if u.startswith("pkt") or u == "packet": return "Pkt"
        if u.startswith("box"): return "Box"
        if u.startswith("pc"):  return "Pcs" if "cs" in u else "Pc"
        if u.startswith("nos"): return "Nos"
        return "Box"

    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        rows = list(csv.reader(f))

    header_idx = None
    for i, row in enumerate(rows):
        joined = " ".join(row).lower()
        if "s. no" in joined or "name (english)" in joined:
            header_idx = i; break

    col_sno = col_name = col_rate = col_disc = col_per = None
    if header_idx is not None:
        for i, h in enumerate([c.strip().lower() for c in rows[header_idx]]):
            if "s. no" in h or h == "s.no":    col_sno = i
            elif "name (english)" in h:        col_name = i
            elif "tamil" in h:                 pass
            elif "discount" in h:              col_disc = i
            elif h.startswith("rate"):         col_rate = i
            elif h.startswith("per"):          col_per = i

    if col_sno  is None: col_sno  = 0
    if col_name is None: col_name = 1
    if col_rate is None: col_rate = 3
    if col_disc is None: col_disc = 4
    if col_per  is None: col_per  = 5

    start = (header_idx + 1) if header_idx is not None else 0

    for ri in range(start, len(rows)):
        row = rows[ri]
        if not row or not any(c.strip() for c in row): continue
        if col_sno >= len(row): continue
        sno_raw = row[col_sno].strip().rstrip(".").strip()
        if not sno_raw.isdigit(): continue
        sno = int(sno_raw)
        if col_name >= len(row): continue
        name = row[col_name].strip()
        if not name: continue
        # Preserve the CSV product name exactly (including Unicode inch symbols
        # such as ½, ¾, ″ and curly quotes). Only remove BOM/outer whitespace.
        name = name.replace("\ufeff", "").strip()
        if not name: continue
        rate = parse_money(row[col_rate]) if col_rate < len(row) else None
        if rate is None or rate <= 0: continue
        disc = parse_money(row[col_disc]) if col_disc < len(row) else None
        if disc is None or disc <= 0: disc = round(rate * 0.20, 2)
        unit = clean_unit(row[col_per]) if col_per < len(row) else "Box"
        pct = 80
        if rate > 0 and disc > 0:
            pct = int(round((1 - disc / rate) * 100))
            pct = max(10, min(90, int(round(pct / 10.0) * 10)))
        products.append({
            "id": f"P{sno:04d}", "sno": sno, "name": name,
            "rate": rate, "discount": disc, "disc_pct": pct,
            "unit": unit, "active": True,
        })

    products.sort(key=lambda p: p["sno"])
    for i, p in enumerate(products, 1):
        p["sno"] = i; p["id"] = f"P{i:04d}"
    return products

# =========================================================
# THEME
# =========================================================
THEMES = {
    "Dark Blue": {
        "BG": "#0f172a", "PANEL": "#1e293b", "INPUT_BG": "#334155",
        "ACCENT": "#f59e0b", "TEXT": "#f1f5f9", "MUTED": "#94a3b8",
        "SUCCESS": "#10b981", "DANGER": "#ef4444", "INFO": "#3b82f6",
        "ROWBG": "#0b1220",
    },
    "Light": {
        "BG": "#f1f5f9", "PANEL": "#ffffff", "INPUT_BG": "#e2e8f0",
        "ACCENT": "#f59e0b", "TEXT": "#0f172a", "MUTED": "#64748b",
        "SUCCESS": "#059669", "DANGER": "#dc2626", "INFO": "#2563eb",
        "ROWBG": "#ffffff",
    },
    "Midnight": {
        "BG": "#020617", "PANEL": "#111827", "INPUT_BG": "#1f2937",
        "ACCENT": "#22d3ee", "TEXT": "#f8fafc", "MUTED": "#94a3b8",
        "SUCCESS": "#34d399", "DANGER": "#fb7185", "INFO": "#60a5fa",
        "ROWBG": "#030712",
    },
}
def _apply_theme(name):
    global BG, PANEL, INPUT_BG, ACCENT, TEXT, MUTED, SUCCESS, DANGER, INFO, ROWBG
    cfg = THEMES.get(name, THEMES["Dark Blue"])
    BG, PANEL, INPUT_BG = cfg["BG"], cfg["PANEL"], cfg["INPUT_BG"]
    ACCENT, TEXT, MUTED = cfg["ACCENT"], cfg["TEXT"], cfg["MUTED"]
    SUCCESS, DANGER, INFO, ROWBG = cfg["SUCCESS"], cfg["DANGER"], cfg["INFO"], cfg["ROWBG"]

_apply_theme(SETTINGS.get("theme", "Dark Blue"))

FONT     = ("Segoe UI", 10)
FONT_B   = ("Segoe UI", 10, "bold")
FONT_H1  = ("Segoe UI", 18, "bold")
FONT_H2  = ("Segoe UI", 13, "bold")
FONT_BIG = ("Segoe UI", 24, "bold")

# =========================================================
# MAIN APP
# =========================================================
class AngelApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Angel Pattasu Kadai — Billing")
        self.geometry("1500x900")
        self.minsize(1100, 700)
        _apply_theme(SETTINGS.get("theme", "Dark Blue"))
        self.configure(bg=BG)
        self.current_view = "show_dashboard"
        try: self.state("zoomed")
        except Exception: pass

        self.cart = []
        self.editing_product_id = None
        self.editing_bill_no = None
        self.editing_bill_original = None

        self.sidebar = tk.Frame(self, bg=PANEL, width=210)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)
        self.content = tk.Frame(self, bg=BG)
        self.content.pack(side="right", fill="both", expand=True)

        self.build_sidebar()
        self.bind_shortcuts()
        self.show_dashboard()

    def build_sidebar(self):
        # Use the supplied round Angel logo in the GUI instead of the emoji/square symbol.
        self.sidebar_logo_img = None
        if Image is not None and os.path.exists(ROUND_LOGO_PATH):
            try:
                img = Image.open(ROUND_LOGO_PATH).convert("RGBA")
                img.thumbnail((118, 118), Image.LANCZOS)
                self.sidebar_logo_img = ImageTk.PhotoImage(img)
                tk.Label(self.sidebar, image=self.sidebar_logo_img, bg=PANEL,
                         bd=0).pack(pady=(12, 2))
                try:
                    self.iconphoto(True, self.sidebar_logo_img)
                except Exception:
                    pass
            except Exception:
                tk.Label(self.sidebar, text="ANGEL", bg=PANEL, fg=ACCENT,
                         font=("Segoe UI", 24, "bold")).pack(pady=(20, 0))
        else:
            tk.Label(self.sidebar, text="ANGEL", bg=PANEL, fg=ACCENT,
                     font=("Segoe UI", 24, "bold")).pack(pady=(20, 0))
        tk.Label(self.sidebar, text="ANGEL PATTASU\nKADAI", bg=PANEL, fg=TEXT,
                 font=FONT_H2, justify="center").pack(pady=(2, 16))
        tk.Frame(self.sidebar, bg=INPUT_BG, height=1).pack(fill="x", padx=14)
        for label, cmd in [
            ("🏠  Dashboard",    self.show_dashboard),
            ("🧾  New Bill",     self.show_billing),
            ("📦  Products",     self.show_products),
            ("📜  Bill History", self.show_history),
            ("📊  Statement",    self.show_statement),
            ("📦  Stock Management", self.show_stock),
            ("⚙️  Settings",     self.show_settings),
        ]:
            self.nav_button(label, cmd)
        tk.Frame(self.sidebar, bg=PANEL).pack(fill="both", expand=True)
        tk.Button(self.sidebar, text="❌  Exit", bg=DANGER, fg="white",
                  font=FONT_B, relief="flat", cursor="hand2",
                  command=self.destroy).pack(fill="x", padx=14, pady=14, ipady=8)

    def nav_button(self, text, cmd):
        btn = tk.Button(self.sidebar, text=f"  {text}", bg=PANEL, fg=TEXT,
                        activebackground=INPUT_BG, activeforeground=ACCENT,
                        font=("Segoe UI", 11), anchor="w", relief="flat",
                        cursor="hand2", padx=14, pady=10, command=cmd)
        btn.pack(fill="x")
        btn.bind("<Enter>", lambda e: btn.config(bg=INPUT_BG))
        btn.bind("<Leave>", lambda e: btn.config(bg=PANEL))
        return btn

    def bind_shortcuts(self):
        self.bind("<F1>", lambda e: self.show_dashboard())
        self.bind("<F2>", lambda e: self.show_billing())
        self.bind("<F3>", lambda e: self.show_products())
        self.bind("<F4>", lambda e: self.show_history())
        self.bind("<F5>", lambda e: self.show_settings())
        self.bind_all("<Control-a>", self._select_all_text)
        self.bind_all("<Control-A>", self._select_all_text)
        self.bind("<F6>", lambda e: self.show_stock())
        self.bind("<F7>", lambda e: self.show_statement())
        self.bind("<F9>", lambda e: self.save_bill())

    def clear_content(self):
        for w in self.content.winfo_children():
            w.destroy()

    # =========================================================
    # DASHBOARD
    # =========================================================
    def show_dashboard(self):
        self.current_view = "show_dashboard"
        self.clear_content()
        head = tk.Frame(self.content, bg=BG)
        head.pack(fill="x", padx=20, pady=(20, 10))
        tk.Label(head, text="🏠  Dashboard", bg=BG, fg=TEXT,
                 font=FONT_H1).pack(side="left")
        tk.Label(head, text=datetime.datetime.now().strftime("%A, %d %B %Y"),
                 bg=BG, fg=MUTED, font=FONT).pack(side="right")

        today = datetime.datetime.now().strftime("%d-%m-%Y")
        today_bills = [b for b in BILLS if b.get("date") == today]
        today_total = sum(_bill_totals(b)["grand_total"] for b in today_bills)
        all_total   = sum(_bill_totals(b)["grand_total"] for b in BILLS)

        stats = tk.Frame(self.content, bg=BG)
        stats.pack(fill="x", padx=20, pady=10)
        for t, v, c in [("🧾 Bills Today", str(len(today_bills)), INFO),
                        ("💰 Sales Today", f"₹ {today_total:,.0f}", SUCCESS),
                        ("📚 Total Bills", str(len(BILLS)), ACCENT),
                        ("💼 Total Sales", f"₹ {all_total:,.0f}", "#8b5cf6"),
                        ("📦 Products", str(len(PRODUCTS)), "#06b6d4"),
                         ("📦 Stock Items", str(len(STOCK)), "#14b8a6")]:
            card = tk.Frame(stats, bg=PANEL)
            card.pack(side="left", expand=True, fill="x", padx=6, ipady=12)
            tk.Frame(card, bg=c, height=4).pack(fill="x")
            tk.Label(card, text=t, bg=PANEL, fg=MUTED,
                     font=("Segoe UI", 9, "bold")).pack(pady=(12, 2))
            tk.Label(card, text=v, bg=PANEL, fg=c,
                     font=("Segoe UI", 20, "bold")).pack(pady=(0, 10))

        qa = tk.Frame(self.content, bg=PANEL)
        qa.pack(fill="x", padx=20, pady=(20, 10))
        tk.Label(qa, text="⚡  Quick Actions", bg=PANEL, fg=ACCENT,
                 font=FONT_H2).pack(anchor="w", padx=16, pady=(12, 6))
        row = tk.Frame(qa, bg=PANEL)
        row.pack(fill="x", padx=16, pady=(0, 16))
        for text, cmd, color in [
            ("🧾  New Bill (F2)",     self.show_billing,   SUCCESS),
            ("📦  Products (F3)",     self.show_products,  INFO),
            ("📜  Bill History (F4)", self.show_history,   ACCENT),
            ("⚙️  Settings (F5)",     self.show_settings,  "#8b5cf6"),
        ]:
            tk.Button(row, text=text, bg=color, fg="white", font=FONT_B,
                      relief="flat", cursor="hand2", padx=20, pady=14,
                      command=cmd).pack(side="left", padx=6, expand=True, fill="x")

        rec = tk.Frame(self.content, bg=PANEL)
        rec.pack(fill="both", expand=True, padx=20, pady=(10, 20))
        tk.Label(rec, text="📋  Recent Bills", bg=PANEL, fg=ACCENT,
                 font=FONT_H2).pack(anchor="w", padx=16, pady=(12, 6))
        self._tree_style()
        cols = ("Bill No", "Date", "Customer", "Total")
        tree = ttk.Treeview(rec, columns=cols, show="headings",
                            style="Angel.Treeview", height=8)
        for c in cols:
            tree.heading(c, text=c); tree.column(c, width=200, anchor="center")
        tree.pack(fill="both", expand=True, padx=16, pady=(0, 16))
        for b in reversed(BILLS[-15:]):
            tree.insert("", "end", values=(
                b.get("bill_no", ""), b.get("date", ""),
                b.get("customer", "") or "-",
                f"₹ {b.get('grand_total',0):,.2f}"))

        # Dashboard Recent Bills: double-click any bill to generate/open its PDF.
        # This is intentionally separate from Bill History double-click, which
        # keeps its existing Edit behavior unchanged.
        tree.bind("<Double-Button-1>", self._dashboard_bill_double_click)
        tree.bind("<Return>", lambda e: self._dashboard_bill_double_click())

    def _dashboard_bill_double_click(self, event=None):
        """Open the selected Dashboard Recent Bill as a PDF."""
        tree = None
        # Find the Recent Bills Treeview on the current Dashboard without
        # changing any existing Dashboard/history controls.
        for widget in self.content.winfo_children():
            for child in widget.winfo_children():
                if isinstance(child, ttk.Treeview):
                    tree = child
                    break
            if tree is not None:
                break

        if tree is None:
            return "break"

        if event is not None:
            row = tree.identify_row(event.y)
            if row:
                tree.selection_set(row)
                tree.focus(row)

        sel = tree.selection()
        if not sel:
            return "break"

        values = tree.item(sel[0], "values")
        if not values:
            return "break"

        bill_no = str(values[0])
        bill = next((b for b in BILLS if str(b.get("bill_no", "")) == bill_no), None)
        if not bill:
            messagebox.showwarning("Bill Not Found", f"Bill {bill_no} was not found.")
            return "break"

        out_pdf = self._history_bill_pdf(bill)
        try:
            # Regenerate from the saved bill data so the Dashboard always opens
            # the current PDF representation, including User 2 GST/overrides and
            # the User 1 composition-tax note.
            generate_a4_invoice(bill, out_pdf)
            if not self._open_pdf_file(out_pdf):
                messagebox.showwarning(
                    "PDF",
                    f"PDF was generated but could not be opened automatically.\n\n{out_pdf}"
                )
        except Exception as e:
            messagebox.showerror("PDF Error", str(e))

        return "break"

    # =========================================================
    # PRODUCTS
    # =========================================================
    def show_products(self):
        self.current_view = "show_products"
        self.clear_content()
        self.editing_product_id = None

        head = tk.Frame(self.content, bg=BG)
        head.pack(fill="x", padx=20, pady=(20, 10))
        tk.Label(head, text="📦  Products", bg=BG, fg=TEXT,
                 font=FONT_H1).pack(side="left")
        self.prod_count_label = tk.Label(head, text=f"{len(PRODUCTS)} items",
                                         bg=BG, fg=MUTED, font=FONT)
        self.prod_count_label.pack(side="right")

        body = tk.Frame(self.content, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        left = tk.Frame(body, bg=BG)
        left.pack(side="left", fill="both", expand=True, padx=(0, 10))
        right = tk.Frame(body, bg=PANEL, width=360)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        top = tk.Frame(left, bg=PANEL)
        top.pack(fill="x", pady=(0, 10))
        self.prod_search_var = tk.StringVar()
        self.prod_search_var.trace_add("write",
                                       lambda *_: self.refresh_product_table())
        e = tk.Entry(top, textvariable=self.prod_search_var,
                     font=("Segoe UI", 12), bg=INPUT_BG, fg=TEXT,
                     insertbackground=TEXT, relief="flat")
        e.pack(side="left", fill="x", expand=True, padx=14, pady=12, ipady=8)
        e.focus_set()
        tk.Button(top, text="📄 Import CSV", bg=ACCENT, fg="#111",
                  font=FONT_B, relief="flat", cursor="hand2",
                  padx=16, pady=8, command=self.import_csv
                  ).pack(side="left", padx=(0, 6))
        tk.Button(top, text="➕ New Product", bg=SUCCESS, fg="white",
                  font=FONT_B, relief="flat", cursor="hand2",
                  padx=16, pady=8, command=self.new_product_form
                  ).pack(side="left", padx=(0, 14))

        self._tree_style()
        cols = ("S.No", "ID", "Name (English)", "Rate ₹", "Disc %",
                "Discount ₹", "Per", "Active")
        wrap = tk.Frame(left, bg=PANEL)
        wrap.pack(fill="both", expand=True)
        self.prod_tree = ttk.Treeview(wrap, columns=cols, show="headings",
                                      style="Angel.Treeview")
        widths = {"S.No":55, "ID":70, "Name (English)":300,
                  "Rate ₹":90, "Disc %":70, "Discount ₹":110,
                  "Per":65, "Active":60}
        for c in cols:
            self.prod_tree.heading(c, text=c)
            self.prod_tree.column(c, width=widths[c],
                                  anchor="w" if c=="Name (English)" else "center")
        self.prod_tree.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(wrap, command=self.prod_tree.yview)
        self.prod_tree.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.prod_tree.bind("<<TreeviewSelect>>", self.on_product_row_select)

        btns = tk.Frame(left, bg=BG)
        btns.pack(fill="x", pady=(10, 0))
        tk.Button(btns, text="🗑 Delete Selected", bg=DANGER, fg="white",
                  font=FONT_B, relief="flat", padx=14, pady=8,
                  command=self.delete_product).pack(side="left", padx=4)
        tk.Button(btns, text="🧹 Clear All Products", bg="#7f1d1d", fg="white",
                  font=FONT_B, relief="flat", padx=14, pady=8,
                  command=self.clear_all_products).pack(side="right", padx=4)

        self._build_product_form(right)
        self.refresh_product_table()
        self.new_product_form()

    def _build_product_form(self, parent):
        hdr = tk.Frame(parent, bg=PANEL)
        hdr.pack(fill="x", padx=14, pady=(14, 6))
        self.form_title = tk.Label(hdr, text="➕  New Product",
                                   bg=PANEL, fg=ACCENT, font=FONT_H2)
        self.form_title.pack(anchor="w")

        body = tk.Frame(parent, bg=PANEL)
        body.pack(fill="both", expand=True, padx=14, pady=(6, 14))

        def row_label(text):
            tk.Label(body, text=text, bg=PANEL, fg=MUTED,
                     font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 0))

        # S.No + ID
        row0 = tk.Frame(body, bg=PANEL); row0.pack(fill="x", pady=(8, 0))
        tk.Label(row0, text="S.No", bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 6))
        self.v_sno = tk.StringVar()
        tk.Entry(row0, textvariable=self.v_sno, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=("Segoe UI", 11),
                 width=8, justify="center").pack(side="left", ipady=6)
        tk.Label(row0, text="ID", bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=(12, 6))
        self.v_id = tk.StringVar()
        tk.Entry(row0, textvariable=self.v_id, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=("Segoe UI", 11),
                 width=12, justify="center").pack(side="left", fill="x",
                                                   expand=True, ipady=6)

        row_label("Name (English)")
        self.v_name = tk.StringVar()
        tk.Entry(body, textvariable=self.v_name, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat",
                 font=("Segoe UI", 11)).pack(fill="x", ipady=6)

        # Rate + Discount %
        row1 = tk.Frame(body, bg=PANEL); row1.pack(fill="x", pady=(8, 0))
        tk.Label(row1, text="Rate (₹)", bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=(0, 6))
        self.v_rate = tk.StringVar()
        tk.Entry(row1, textvariable=self.v_rate, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat",
                 font=("Segoe UI", 11), width=10, justify="right"
                 ).pack(side="left", ipady=6)
        tk.Label(row1, text="Disc %", bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=(12, 6))
        self.v_disc_pct = tk.StringVar(value=str(SETTINGS.get("default_discount_pct", 80)))
        ttk.Combobox(row1, textvariable=self.v_disc_pct,
                     values=[str(x) for x in range(10, 100, 10)],
                     width=5, state="readonly",
                     font=("Segoe UI", 11)).pack(side="left", ipady=4)

        row_label("Discount Price (₹)  — editable / auto calculated")
        self.v_disc = tk.StringVar(value="0.00")
        tk.Entry(body, textvariable=self.v_disc, bg=ROWBG, fg=SUCCESS,
                 insertbackground=SUCCESS, relief="flat",
                 font=("Segoe UI", 13, "bold"),
                 justify="right"
                 ).pack(fill="x", ipady=8)

        # Keep Rate and Discount Price bidirectionally linked.
        # Rate/Disc% changes update Discount Price; editing Discount Price
        # updates Rate automatically (e.g. 80% discount: ₹200 -> ₹1000).
        self._product_calc_lock = False
        self.v_rate.trace_add("write", lambda *_: self._recalc_discount())
        self.v_disc_pct.trace_add("write", lambda *_: self._recalc_discount())
        self.v_disc.trace_add("write", lambda *_: self._recalc_rate_from_discount())

        row_label("Per (Box / Pkt / Pc / Nos)")
        self.v_unit = tk.StringVar(value="Box")
        ttk.Combobox(body, textvariable=self.v_unit,
                     values=["Box", "Pkt", "Pc", "Pcs", "Nos"],
                     state="readonly", font=("Segoe UI", 11)
                     ).pack(fill="x", ipady=4)

        self.v_active = tk.BooleanVar(value=True)
        tk.Checkbutton(body, text="Active", variable=self.v_active,
                       bg=PANEL, fg=TEXT, selectcolor=PANEL,
                       activebackground=PANEL, activeforeground=TEXT,
                       font=FONT).pack(anchor="w", pady=(10, 0))

        btn_row = tk.Frame(parent, bg=PANEL)
        btn_row.pack(fill="x", padx=14, pady=(0, 14))
        tk.Button(btn_row, text="💾 Save Product", bg=SUCCESS, fg="white",
                  font=FONT_B, relief="flat", cursor="hand2",
                  padx=14, pady=10, command=self.save_product_form
                  ).pack(fill="x", pady=(0, 6))
        tk.Button(btn_row, text="↺ Clear / New", bg=INFO, fg="white",
                  font=FONT_B, relief="flat", cursor="hand2",
                  padx=14, pady=8, command=self.new_product_form
                  ).pack(fill="x")

    def _recalc_discount(self):
        if getattr(self, "_product_calc_lock", False):
            return
        try: rate = float(self.v_rate.get() or 0)
        except ValueError: rate = 0.0
        try: pct = float(self.v_disc_pct.get() or 0)
        except ValueError: pct = 0.0
        disc = round(rate * (1 - pct / 100.0), 2)
        self._product_calc_lock = True
        try:
            self.v_disc.set(f"{disc:.2f}")
        finally:
            self._product_calc_lock = False

    def _recalc_rate_from_discount(self):
        if getattr(self, "_product_calc_lock", False):
            return
        try:
            disc = float(self.v_disc.get() or 0)
        except ValueError:
            return
        try:
            pct = float(self.v_disc_pct.get() or 0)
        except ValueError:
            pct = 80.0
        multiplier = 1.0 - (pct / 100.0)
        if multiplier <= 0:
            return
        rate = round(disc / multiplier, 2)
        self._product_calc_lock = True
        try:
            self.v_rate.set(f"{rate:.2f}")
        finally:
            self._product_calc_lock = False

    def new_product_form(self):
        self.editing_product_id = None
        next_sno = max((p.get("sno", 0) for p in PRODUCTS), default=0) + 1
        self.form_title.config(text="➕  New Product")
        self.v_sno.set(str(next_sno))
        self.v_id.set(f"P{next_sno:04d}")
        self.v_name.set("")
        self.v_rate.set("")
        self.v_disc_pct.set(str(SETTINGS.get("default_discount_pct", 80)))
        self.v_disc.set("0.00")
        self.v_unit.set("Box")
        self.v_active.set(True)

    def on_product_row_select(self, event=None):
        sel = self.prod_tree.selection()
        if not sel: return
        prod = next((p for p in PRODUCTS if p["id"] == sel[0]), None)
        if not prod: return

        self.editing_product_id = prod["id"]
        self.form_title.config(text=f"✏️  Editing: {prod['name'][:26]}")
        self.v_sno.set(str(prod.get("sno", "")))
        self.v_id.set(prod["id"])
        self.v_name.set(prod["name"])
        self.v_rate.set(f"{prod['rate']:.2f}")
        rate = prod["rate"]; disc = prod.get("discount", rate)
        if "disc_pct" in prod:
            pct = int(prod["disc_pct"])
        elif rate > 0:
            pct = int(round((1 - disc / rate) * 100))
            pct = max(10, min(90, int(round(pct / 10.0) * 10)))
        else:
            pct = 80
        self.v_disc_pct.set(str(pct))
        self.v_disc.set(f"{disc:.2f}")
        self.v_unit.set(prod.get("unit", "Box"))
        self.v_active.set(prod.get("active", True))

    def save_product_form(self):
        try:
            sno  = int(self.v_sno.get().strip())
            rate = float(self.v_rate.get().strip())
        except ValueError:
            messagebox.showerror("Invalid",
                "S.No must be an integer and Rate must be a number.")
            return
        try: pct = float(self.v_disc_pct.get() or 80)
        except ValueError: pct = 80
        try:
            disc = float(self.v_disc.get().strip())
        except ValueError:
            messagebox.showerror("Invalid", "Discount Price must be a number.")
            return
        if disc < 0:
            messagebox.showerror("Invalid", "Discount Price cannot be negative.")
            return
        if rate >= 0 and disc > rate:
            messagebox.showerror("Invalid", "Discount Price cannot be greater than Rate.")
            return
        disc = round(disc, 2)

        data = {
            "id": self.v_id.get().strip(), "sno": sno,
            "name": self.v_name.get().strip(),
            "rate": rate, "discount": disc, "disc_pct": pct,
            "unit": self.v_unit.get().strip() or "Box",
            "active": bool(self.v_active.get()),
        }
        if not data["id"] or not data["name"]:
            messagebox.showerror("Invalid", "ID and Name are required.")
            return

        if self.editing_product_id:
            for i, p in enumerate(PRODUCTS):
                if p["id"] == self.editing_product_id:
                    if data["id"] != self.editing_product_id and \
                       any(q["id"] == data["id"] for q in PRODUCTS):
                        messagebox.showerror("Duplicate",
                            f"Product ID {data['id']} already exists.")
                        return
                    PRODUCTS[i] = data; break
        else:
            if any(p["id"] == data["id"] for p in PRODUCTS):
                messagebox.showerror("Duplicate",
                    f"Product ID {data['id']} already exists.")
                return
            PRODUCTS.append(data)

        save_json(PRODUCTS_FILE, PRODUCTS)
        self.refresh_product_table()
        self.prod_count_label.config(text=f"{len(PRODUCTS)} items")
        self.new_product_form()

    def refresh_product_table(self):
        for r in self.prod_tree.get_children():
            self.prod_tree.delete(r)
        q = self.prod_search_var.get().lower().strip()
        for p in sorted(PRODUCTS, key=lambda x: x.get("sno", 99999)):
            hay = f"{p.get('sno','')} {p['id']} {p['name']}".lower()
            if q and q not in hay: continue
            rate = p["rate"]; disc = p.get("discount", rate)
            pct = p.get("disc_pct",
                        int(round((1 - disc / rate) * 100)) if rate > 0 else 0)
            self.prod_tree.insert("", "end", iid=p["id"], values=(
                p.get("sno", ""), p["id"], p["name"],
                f"{rate:.2f}", f"{pct}%", f"{disc:.2f}",
                p.get("unit", "Box"),
                "✓" if p.get("active", True) else "✗"))

    def import_csv(self):
        path = filedialog.askopenfilename(
            title="Select price list CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")])
        if not path: return
        try: items = import_from_csv(path)
        except Exception as e:
            messagebox.showerror("Import failed", f"{e}"); return
        if not items:
            messagebox.showwarning("No rows found",
                "Could not read any product rows.")
            return
        if PRODUCTS:
            replace = messagebox.askyesno("Existing products",
                f"{len(PRODUCTS)} products already exist.\n\n"
                f"Replace them with {len(items)} imported items?\n\n"
                "Yes = Replace all\nNo  = Append")
        else:
            replace = True
        if replace: PRODUCTS.clear()
        PRODUCTS.extend(items)
        for i, p in enumerate(sorted(PRODUCTS, key=lambda x: x.get("sno", 0)), 1):
            p["sno"] = i
        save_json(PRODUCTS_FILE, PRODUCTS)
        self.refresh_product_table()
        self.prod_count_label.config(text=f"{len(PRODUCTS)} items")
        messagebox.showinfo("Import successful",
            f"Imported {len(items)} products.")

    def delete_product(self):
        sel = self.prod_tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a product row first.")
            return
        prod = next((p for p in PRODUCTS if p["id"] == sel[0]), None)
        if not prod: return
        if messagebox.askyesno("Delete", f"Delete product:\n{prod['name']}?"):
            PRODUCTS.remove(prod)
            save_json(PRODUCTS_FILE, PRODUCTS)
            self.refresh_product_table()
            self.prod_count_label.config(text=f"{len(PRODUCTS)} items")
            self.new_product_form()

    def clear_all_products(self):
        if not PRODUCTS: return
        if messagebox.askyesno("Clear All",
                              f"Delete ALL {len(PRODUCTS)} products?"):
            PRODUCTS.clear()
            save_json(PRODUCTS_FILE, PRODUCTS)
            self.refresh_product_table()
            self.prod_count_label.config(text="0 items")
            self.new_product_form()

    # =========================================================
    # BILLING  (fast inline-editable cart)
    # =========================================================
    def show_billing(self, edit_bill=None):
        self.current_view = "show_billing"
        self.clear_content()
        self._cancel_cart_editor()

        # History -> Edit always comes back into this same New Bill GUI.
        if edit_bill is not None:
            self.editing_bill_no = str(edit_bill.get("bill_no", ""))
            self.editing_bill_original = copy.deepcopy(edit_bill)
            self.cart = [copy.deepcopy(it) for it in edit_bill.get("items", [])]
        else:
            self.editing_bill_no = None
            self.editing_bill_original = None
            self.cart = []

        self.cart_editor = None
        self.cart_editor_item = None
        self.cart_editor_col = None
        self.cart_search_var = tk.StringVar(value="")

        # User 2 Total GST: automatic by default; becomes manual only when
        # the user actually edits the Total GST field.
        self.total_gst_manual = False
        self.total_gst_override = tk.StringVar(value="0.00")

        # Document type is selected per bill and saved with the bill.
        # Existing bills without this field remain Bill Of Supply.
        self.bill_type = tk.StringVar(
            value=str(edit_bill.get("bill_type", "Bill Of Supply") or "Bill Of Supply")
            if edit_bill is not None else "Bill Of Supply"
        )
        if self.bill_type.get() not in ("Bill Of Supply", "Tax Invoice", "Estimate"):
            self.bill_type.set("Bill Of Supply")

        head = tk.Frame(self.content, bg=BG)
        head.pack(fill="x", padx=20, pady=(20, 10))
        title = "✏  Edit Bill" if edit_bill is not None else "🧾  New Bill"
        tk.Label(head, text=title, bg=BG, fg=TEXT,
                 font=FONT_H1).pack(side="left")

        # Bill Type + editable Bill Number are always visible at the top.
        top_bill = tk.Frame(head, bg=BG)
        top_bill.pack(side="right")

        tk.Label(top_bill, text="Bill Type:", bg=BG, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(0, 5))
        self.bill_type_combo = ttk.Combobox(
            top_bill, textvariable=self.bill_type,
            values=["Bill Of Supply", "Tax Invoice", "Estimate"],
            state="readonly", width=16, font=("Segoe UI", 10, "bold")
        )
        self.bill_type_combo.pack(side="left", padx=(0, 14), ipady=3)
        self.bill_type_combo.bind("<<ComboboxSelected>>", self._on_bill_type_changed)

        tk.Label(top_bill, text="Bill No:", bg=BG, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(0, 5))
        self.v_bill_no = tk.StringVar(
            value=str(edit_bill.get("bill_no", "")) if edit_bill is not None
            else self._next_bill_no(self.bill_type.get())
        )
        self.bill_no_entry = tk.Entry(
            top_bill, textvariable=self.v_bill_no,
            bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
            relief="flat", font=("Segoe UI", 12, "bold"),
            width=15, justify="center"
        )
        self.bill_no_entry.pack(side="left", ipady=5)

        # Editable bill date. Existing bills retain their saved date;
        # new bills start with today's date.
        initial_bill_date = (
            str(edit_bill.get("date", "") or "").strip()
            if edit_bill is not None else ""
        )
        if not initial_bill_date:
            initial_bill_date = datetime.datetime.now().strftime("%d-%m-%Y")
        self.bill_date = tk.StringVar(value=initial_bill_date)

        tk.Label(top_bill, text="Date:", bg=BG, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(14, 5))
        self.bill_date_entry = tk.Entry(
            top_bill, textvariable=self.bill_date,
            bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
            relief="flat", font=("Segoe UI", 11, "bold"),
            width=12, justify="center"
        )
        self.bill_date_entry.pack(side="left", ipady=5)

        body = tk.Frame(self.content, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        left = tk.Frame(body, bg=PANEL, width=400)
        left.pack(side="left", fill="y", padx=(0, 10))
        left.pack_propagate(False)
        right = tk.Frame(body, bg=PANEL)
        right.pack(side="right", fill="both", expand=True)

        self._billing_left(left)
        self._billing_right(right)

        # Billing mode: User 1 remains the existing bill. User 2 enables GST only.
        if not hasattr(self, "billing_user"):
            self.billing_user = tk.StringVar(value="User 1")
        if not hasattr(self, "cgst_pct"):
            self.cgst_pct = tk.StringVar(value="9")
        if not hasattr(self, "sgst_pct"):
            self.sgst_pct = tk.StringVar(value="9")

        if edit_bill is not None:
            self.cust_name.set(edit_bill.get("customer", ""))
            self.cust_mob.set(edit_bill.get("mobile", ""))
            # Defensive creation keeps old bills compatible even if address was absent.
            if hasattr(self, "cust_address"):
                self.cust_address.delete("1.0", tk.END)
                self.cust_address.insert("1.0", edit_bill.get("address", "") or "")
            self.packing_charge.set(f"{float(edit_bill.get('packing_charge', 0) or 0):.2f}")
            self.billing_user.set(edit_bill.get("billing_user", "User 1") or "User 1")
            self.cgst_pct.set(str(edit_bill.get("cgst_pct", 9) or 9))
            self.sgst_pct.set(str(edit_bill.get("sgst_pct", 9) or 9))
            self.total_gst_manual = bool(edit_bill.get("total_gst_manual", False))
            self.total_gst_override.set(
                f"{_safe_number(edit_bill.get('total_gst_override', edit_bill.get('total_gst', 0))):.2f}"
            )
            self._update_billing_user_mode()
            self.amount_paid.set(f"{_safe_number(edit_bill.get('amount_paid', 0)):.2f}")
            self._update_payment_from_amount()
        else:
            if hasattr(self, "cust_address"):
                self.cust_address.delete("1.0", tk.END)
            self.packing_charge.set("0.00")
            self.billing_user.set("User 1")
            self.cgst_pct.set("9")
            self.sgst_pct.set("9")
            self.total_gst_manual = False
            self.total_gst_override.set("0.00")
            self._update_billing_user_mode()
            # New bills always start as Paid. User can change the dropdown
            # to Not Paid or Partial when required.
            self.payment_status.set("Paid")
            self.amount_paid.set("0.00")

        self.refresh_cart()
        self.search_entry.focus_set()

    def _bill_type_prefix(self, bill_type):
        return BILL_TYPE_CONFIG.get(
            str(bill_type or "").strip(),
            BILL_TYPE_CONFIG["Bill Of Supply"]
        )["prefix"]

    def _bill_type_pdf_dir(self, bill_type):
        return BILL_TYPE_PDF_DIRS.get(
            str(bill_type or "").strip(),
            BILL_TYPE_PDF_DIRS["Bill Of Supply"]
        )

    def _next_user3_bill_no(self):
        """Return User 3 transport Bill Of Supply numbers as plain 001, 002, 003..."""
        counters = SETTINGS.setdefault("bill_counters", {})
        try:
            next_num = int(counters.get("USER3", 1) or 1)
        except (TypeError, ValueError):
            next_num = 1
        max_existing = 0
        for saved_bill in BILLS:
            if str(saved_bill.get("billing_user", "")).strip() != "User 3":
                continue
            saved_no = str(saved_bill.get("bill_no", "") or "").strip()
            if re.fullmatch(r"\d{3}", saved_no):
                max_existing = max(max_existing, int(saved_no))
        next_num = max(next_num, max_existing + 1)
        counters["USER3"] = next_num
        return f"{next_num:03d}"

    def _next_bill_no(self, bill_type=None):
        """Return the next number independently for the selected bill type."""
        bill_type = str(
            bill_type or getattr(self, "bill_type", tk.StringVar(
                value="Bill Of Supply"
            )).get() or "Bill Of Supply"
        ).strip()
        if bill_type not in BILL_TYPE_CONFIG:
            bill_type = "Bill Of Supply"

        prefix = self._bill_type_prefix(bill_type)

        # Prefer the per-type counter, but also inspect existing bills so
        # older data can never accidentally reuse an already-used number.
        counters = SETTINGS.setdefault("bill_counters", {})
        stored = counters.get(prefix, 1)
        try:
            stored = max(1, int(stored))
        except (TypeError, ValueError):
            stored = 1

        max_existing = 0
        pattern = re.compile(r"^" + re.escape(prefix) + r"-(\\d+)$")
        for saved_bill in BILLS:
            m = pattern.match(str(saved_bill.get("bill_no", "")).strip())
            if m:
                try:
                    max_existing = max(max_existing, int(m.group(1)))
                except ValueError:
                    pass

        n = max(stored, max_existing + 1)
        return f"{prefix}-{n:05d}"

    def _bill_pdf_path(self, bill, new_file=False):
        """Build a collision-safe PDF path containing bill type, customer and timestamp."""
        bill_type = str(
            bill.get("bill_type", "Bill Of Supply") or "Bill Of Supply"
        ).strip()
        if bill_type not in BILL_TYPE_CONFIG:
            bill_type = "Bill Of Supply"

        bill_no = str(bill.get("bill_no", "bill") or "bill").strip()
        safe_no = re.sub(r"[^A-Za-z0-9_-]+", "_", bill_no).strip("_") or "bill"

        customer = str(bill.get("customer", "") or "").strip()
        safe_customer = re.sub(r"[^A-Za-z0-9_-]+", "_", customer).strip("_")
        if not safe_customer:
            safe_customer = "Customer"

        if new_file:
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        else:
            stamp = str(bill.get("pdf_timestamp", "") or "").strip()
            if not stamp:
                stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")

        return os.path.join(
            self._bill_type_pdf_dir(bill_type),
            f"{safe_no}_{safe_customer}_{stamp}.pdf"
        )

    def _on_bill_type_changed(self, event=None):
        """For a new bill, selecting a type also switches to that type's number."""
        if getattr(self, "editing_bill_no", None):
            return
        if getattr(self, "billing_user", tk.StringVar(value="User 1")).get().strip() == "User 3":
            self.bill_type.set("Bill Of Supply")
            if hasattr(self, "v_bill_no"):
                self.v_bill_no.set(self._next_user3_bill_no())
            return
        if hasattr(self, "v_bill_no"):
            self.v_bill_no.set(self._next_bill_no(self.bill_type.get()))

    # ---------------------------------------------------------
    # Product search
    # ---------------------------------------------------------
    def _billing_left(self, parent):
        tk.Label(parent, text="🔍  Search Product", bg=PANEL, fg=ACCENT,
                 font=FONT_H2).pack(anchor="w", padx=12, pady=(12, 4))

        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.filter_products())
        self.search_entry = tk.Entry(
            parent, textvariable=self.search_var, font=("Segoe UI", 13),
            bg=INPUT_BG, fg=TEXT, insertbackground=TEXT, relief="flat"
        )
        self.search_entry.pack(fill="x", padx=12, ipady=10)
        self.search_entry.bind("<Return>", self._add_first_match)
        self.search_entry.bind("<Tab>", self._focus_product_list)
        self.search_entry.bind("<Down>", self._focus_product_list)
        self.search_entry.bind("<Escape>", lambda e: self.search_var.set(""))

        tk.Label(
            parent,
            text="Type to search  •  ↓ / Tab = select result  •  Tab = add selected  •  Enter = add first",
            bg=PANEL, fg=MUTED, font=("Segoe UI", 9)
        ).pack(anchor="w", padx=12, pady=(2, 8))

        wrap = tk.Frame(parent, bg=PANEL)
        self.product_results_frame = wrap
        # Keep the billing screen clean: the product results area appears only
        # after the user types a search term.
        self.prod_list = tk.Listbox(
            wrap, font=("Segoe UI", 11), bg=INPUT_BG, fg=TEXT,
            selectbackground=ACCENT, selectforeground="#111",
            relief="flat", activestyle="none"
        )
        sb = tk.Scrollbar(wrap, command=self.prod_list.yview)
        self.prod_list.config(yscrollcommand=sb.set)
        self.prod_list.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.prod_list.bind("<Double-Button-1>", self._add_selected)
        self.prod_list.bind("<Return>", self._add_selected)
        self.prod_list.bind("<Tab>", self._add_selected_tab)
        self.prod_list.bind("<Up>", self._product_list_up)
        self.prod_list.bind("<Down>", self._product_list_down)

        self.manual_item_button = tk.Button(
            parent, text="✚  Manual Item", bg=INFO, fg="white",
            font=FONT_B, relief="flat", cursor="hand2",
            command=self._add_manual_item
        )
        self.manual_item_button.pack(fill="x", padx=12, pady=(0, 12), ipady=7)

        self.filtered_products = []
        self.filter_products()

    def _focus_product_list(self, event=None):
        if self.filtered_products:
            if hasattr(self, "product_results_frame"):
                self.product_results_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8), before=self.manual_item_button)
            self.prod_list.focus_set()
            if not self.prod_list.curselection():
                self.prod_list.selection_set(0)
                self.prod_list.activate(0)
        return "break"

    def _product_list_up(self, event=None):
        if not self.filtered_products:
            return "break"
        sel = self.prod_list.curselection()
        idx = max(0, (sel[0] - 1) if sel else 0)
        self.prod_list.selection_clear(0, tk.END)
        self.prod_list.selection_set(idx)
        self.prod_list.activate(idx)
        self.prod_list.see(idx)
        return "break"

    def _product_list_down(self, event=None):
        if not self.filtered_products:
            return "break"
        sel = self.prod_list.curselection()
        idx = min(len(self.filtered_products) - 1, (sel[0] + 1) if sel else 0)
        self.prod_list.selection_clear(0, tk.END)
        self.prod_list.selection_set(idx)
        self.prod_list.activate(idx)
        self.prod_list.see(idx)
        return "break"

    def _add_selected_tab(self, event=None):
        self._add_selected()
        return "break"

    def filter_products(self):
        """Flexible product search.

        Search is intentionally forgiving so users do not need to type the
        exact product name or remember every part of it.  Examples:
          30cm electric      -> 30cm Electric Sparklers
          30 electric        -> 30cm Electric Sparklers
          electric 30        -> 30cm Electric Sparklers
          30c elec           -> 30cm Electric Sparklers
          spark              -> Sparklers

        Every search word may match part of a product word, and search words
        can be entered in any order. A small typo is also tolerated.
        """
        from difflib import SequenceMatcher

        raw_q = self.search_var.get().strip().lower()
        active = [p for p in PRODUCTS if p.get("active", True)]

        if not raw_q:
            self.filtered_products = []
            self.prod_list.delete(0, tk.END)
            if hasattr(self, "product_results_frame"):
                self.product_results_frame.pack_forget()
            return

        def normalize_words(value):
            s = str(value or "").lower()
            s = re.sub(r"[^0-9a-zA-Z]+", " ", s)
            return [w for w in s.split() if w]

        def compact(value):
            return re.sub(r"[^0-9a-zA-Z]+", "", str(value or "").lower())

        query_words = normalize_words(raw_q)
        query_compact = compact(raw_q)

        scored = []
        for p in active:
            name = str(p.get("name", ""))
            pid = str(p.get("id", ""))
            name_words = normalize_words(name)
            name_compact = compact(name)
            id_words = normalize_words(pid)
            id_compact = compact(pid)

            # Exact full-string matches remain strongest.
            score = 0
            if raw_q in name.lower():
                score += 1000
            if query_compact and query_compact in name_compact:
                score += 900
            if raw_q in pid.lower() or (query_compact and query_compact in id_compact):
                score += 850

            all_words_match = True
            matched_words = 0
            fuzzy_matches = 0

            for qw in query_words:
                # A query word may be a prefix/part of a product word.
                direct = any(qw in nw or nw in qw for nw in name_words)
                if direct:
                    matched_words += 1
                    score += 100
                    continue

                # Also allow a small spelling mistake for longer words.
                best_ratio = 0.0
                for nw in name_words:
                    ratio = SequenceMatcher(None, qw, nw).ratio()
                    prefix_ratio = SequenceMatcher(None, qw, nw[:len(qw)]).ratio()
                    best_ratio = max(best_ratio, ratio, prefix_ratio)

                if len(qw) >= 4 and best_ratio >= 0.78:
                    matched_words += 1
                    fuzzy_matches += 1
                    score += 45
                else:
                    # Try the product ID as a fallback.
                    id_direct = any(qw in iw or iw in qw for iw in id_words)
                    if id_direct:
                        matched_words += 1
                        score += 60
                    else:
                        all_words_match = False
                        break

            if all_words_match and query_words:
                score += matched_words * 20
                if fuzzy_matches:
                    score -= fuzzy_matches * 5
                scored.append((score, name.lower(), p))

        scored.sort(key=lambda item: (-item[0], item[1]))
        self.filtered_products = [item[2] for item in scored]

        self.prod_list.delete(0, tk.END)
        for p in self.filtered_products:
            price = p.get("discount", p.get("rate", 0))
            self.prod_list.insert(
                tk.END,
                f"  {p.get('name', '')}  —  ₹{float(price):.0f}/{p.get('unit', 'Box')}"
            )

        if hasattr(self, "product_results_frame"):
            self.product_results_frame.pack(fill="both", expand=True, padx=12, pady=(0, 8), before=self.manual_item_button)

        if self.filtered_products:
            self.prod_list.selection_clear(0, tk.END)
            self.prod_list.selection_set(0)
            self.prod_list.activate(0)
            self.prod_list.see(0)

    def _add_first_match(self, event=None):
        q = self.search_var.get().strip()
        if self.filtered_products:
            self._add_product(self.filtered_products[0])
        elif q:
            # Product not in master list: add only to this bill.
            self._add_manual_item(name=q)
        return "break"

    def _add_selected(self, event=None):
        sel = self.prod_list.curselection()
        if not sel:
            if self.filtered_products:
                self._add_product(self.filtered_products[0])
            elif self.search_var.get().strip():
                self._add_manual_item(name=self.search_var.get().strip())
            return "break"
        self._add_product(self.filtered_products[sel[0]])
        return "break"

    # ---------------------------------------------------------
    # Add product -> immediately edit Qty
    # ---------------------------------------------------------
    def _add_product(self, prod):
        orig = float(prod.get("rate", 0) or 0)
        pct = float(prod.get(
            "disc_pct", SETTINGS.get("default_discount_pct", 80)
        ) or 0)
        rate = float(prod.get(
            "discount", round(orig * (1 - pct / 100.0), 2)
        ) or 0)

        for idx, it in enumerate(self.cart):
            if (it.get("id") == prod.get("id")
                    and float(it.get("rate", 0)) == rate
                    and float(it.get("orig_rate", 0)) == orig):
                it["qty"] = float(it.get("qty", 0)) + 1
                self.search_var.set("")
                if hasattr(self, "cart_search_var"):
                    self.cart_search_var.set("")
                self.refresh_cart()
                self._start_qty_edit(idx)
                return

        self.cart.append({
            "id": prod.get("id", ""),
            "name": prod.get("name", ""),
            "qty": 1.0,
            "orig_rate": orig,
            "disc_pct": pct,
            "rate": rate,
            "unit": prod.get("unit", "Box"),
        })
        idx = len(self.cart) - 1
        self.search_var.set("")
        if hasattr(self, "cart_search_var"):
            self.cart_search_var.set("")
        self.refresh_cart()
        self._start_qty_edit(idx)

    def _add_manual_item(self, name=""):
        """Bill-only item; never added to the Products master list."""
        self._cancel_cart_editor()
        self.cart.append({
            "id": "",
            "name": str(name).strip(),
            "qty": 1.0,
            "orig_rate": 0.0,
            "disc_pct": 80.0,
            "rate": 0.0,
            "unit": "Box",
            "manual": True,
        })
        idx = len(self.cart) - 1
        self.search_var.set("")
        if hasattr(self, "cart_search_var"):
            self.cart_search_var.set("")
        self.refresh_cart()
        if name:
            self._start_qty_edit(idx)
        else:
            self._start_cart_edit(idx, "Product")

    def _start_qty_edit(self, idx):
        self._start_cart_edit(idx, "Qty")

    # ---------------------------------------------------------
    # Billing cart - ALL fields inline editable, no popup
    # ---------------------------------------------------------
    def _billing_right(self, parent):
        """Build a compact, resize-friendly billing panel.

        The cart is the only expandable section. Customer/payment/GST controls,
        totals and action buttons stay visible even when User 2 is selected.
        Existing billing variables and callbacks are intentionally preserved.
        """
        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        top = tk.Frame(parent, bg=PANEL)
        top.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 4))
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=1)

        # Customer / mobile / billing user
        cust = tk.Frame(top, bg=PANEL)
        cust.pack(fill="x", pady=(0, 5))
        cust.columnconfigure(1, weight=1)
        cust.columnconfigure(3, weight=0)
        cust.columnconfigure(5, weight=0)

        tk.Label(cust, text="Customer:", bg=PANEL, fg=MUTED,
                 font=FONT_B).grid(row=0, column=0, sticky="w", padx=(2, 4))
        self.cust_name = tk.StringVar()
        tk.Entry(cust, textvariable=self.cust_name, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT
                 ).grid(row=0, column=1, sticky="ew", padx=(0, 10), ipady=5)

        tk.Label(cust, text="Mobile:", bg=PANEL, fg=MUTED,
                 font=FONT_B).grid(row=0, column=2, sticky="w", padx=(0, 4))
        self.cust_mob = tk.StringVar()
        tk.Entry(cust, textvariable=self.cust_mob, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT, width=14
                 ).grid(row=0, column=3, sticky="ew", padx=(0, 10), ipady=5)

        tk.Label(cust, text="Billing User:", bg=PANEL, fg=MUTED,
                 font=FONT_B).grid(row=0, column=4, sticky="w", padx=(0, 4))
        self.billing_user = tk.StringVar(value="User 1")
        self.billing_user_combo = ttk.Combobox(
            cust, textvariable=self.billing_user,
            values=["User 1", "User 2", "User 3", "User 4"], state="readonly",
            width=9, font=("Segoe UI", 10, "bold")
        )
        self.billing_user_combo.grid(row=0, column=5, sticky="e", ipady=3)
        self.billing_user_combo.bind(
            "<<ComboboxSelected>>",
            lambda e: self._on_billing_user_changed()
        )

        # Address
        addr_row = tk.Frame(top, bg=PANEL)
        addr_row.pack(fill="x", pady=(0, 5))
        tk.Label(addr_row, text="Address:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", anchor="n", padx=(2, 8))
        self.cust_address = tk.Text(
            addr_row, height=2, wrap="word",
            bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
            relief="flat", font=FONT, undo=True
        )
        self.cust_address.pack(side="left", fill="x", expand=True, ipady=2)
        self.cust_address.bind("<Control-a>", lambda e: self._select_all_text(e))
        self.cust_address.bind("<Control-A>", lambda e: self._select_all_text(e))

        # Payment / amount paid / cart search
        pay = tk.Frame(top, bg=PANEL)
        pay.pack(fill="x", pady=(0, 5))

        tk.Label(pay, text="Payment:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left")
        self.payment_status = tk.StringVar(value="Paid")
        self.payment_combo = ttk.Combobox(
            pay, textvariable=self.payment_status,
            values=["Paid", "Partial", "Not Paid"],
            state="readonly", width=11,
            font=("Segoe UI", 10, "bold")
        )
        self.payment_combo.pack(side="left", padx=5, ipady=3)
        self.payment_combo.bind(
            "<<ComboboxSelected>>",
            lambda e: self._apply_payment_status()
        )

        tk.Label(pay, text="Amount Paid ₹:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(8, 4))
        self.amount_paid = tk.StringVar(value="0.00")
        self.amount_paid_entry = tk.Entry(
            pay, textvariable=self.amount_paid,
            bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
            relief="flat", font=FONT_B, width=12, justify="right"
        )
        self.amount_paid_entry.pack(side="left", ipady=5)
        self.amount_paid_entry.bind(
            "<Return>", lambda e: self._update_payment_from_amount()
        )
        self.amount_paid_entry.bind(
            "<FocusOut>", lambda e: self._update_payment_from_amount()
        )

        # Live balance indicator for Partial payments.
        self.payment_balance_label = tk.Label(
            pay, text="Balance: ₹ 0.00", bg=PANEL, fg=ACCENT,
            font=("Segoe UI", 9, "bold")
        )
        self.payment_balance_label.pack(side="left", padx=(8, 12))

        tk.Label(pay, text="Search Item:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(4, 4))
        self.cart_search_var = tk.StringVar(value="")
        self.cart_search_entry = tk.Entry(
            pay, textvariable=self.cart_search_var,
            bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
            relief="flat", font=FONT_B, width=20
        )
        self.cart_search_entry.pack(side="left", ipady=5, fill="x", expand=True)
        self.cart_search_entry.bind("<KeyRelease>", self._search_cart_rows)
        self.cart_search_entry.bind("<Return>", self._focus_searched_cart_row)
        self.cart_search_entry.bind("<Escape>", self._clear_cart_search)

        # User 2 GST controls. This frame is shown/hidden by
        # _update_billing_user_mode(). It never consumes cart height when hidden.
        self.gst_frame = tk.Frame(top, bg=PANEL)
        gst = self.gst_frame
        tk.Label(gst, text="CGST %:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left")
        self.cgst_pct = tk.StringVar(value="9")
        self.cgst_entry = tk.Entry(
            gst, textvariable=self.cgst_pct, bg=INPUT_BG, fg=TEXT,
            insertbackground=TEXT, relief="flat", font=FONT_B,
            width=7, justify="right"
        )
        self.cgst_entry.pack(side="left", padx=(5, 12), ipady=3)
        self.cgst_entry.bind("<KeyRelease>", lambda e: self.refresh_cart())
        self.cgst_entry.bind("<Return>", lambda e: self.refresh_cart())

        tk.Label(gst, text="SGST %:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left")
        self.sgst_pct = tk.StringVar(value="9")
        self.sgst_entry = tk.Entry(
            gst, textvariable=self.sgst_pct, bg=INPUT_BG, fg=TEXT,
            insertbackground=TEXT, relief="flat", font=FONT_B,
            width=7, justify="right"
        )
        self.sgst_entry.pack(side="left", padx=(5, 12), ipady=3)
        self.sgst_entry.bind("<KeyRelease>", lambda e: self.refresh_cart())
        self.sgst_entry.bind("<Return>", lambda e: self.refresh_cart())

        tk.Label(gst, text="Total GST ₹:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left")
        self.total_gst_entry = tk.Entry(
            gst, textvariable=self.total_gst_override,
            bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
            relief="flat", font=FONT_B, width=12, justify="right"
        )
        self.total_gst_entry.pack(side="left", padx=(5, 4), ipady=3)
        self.total_gst_entry.bind("<KeyPress>", self._mark_total_gst_manual)
        self.total_gst_entry.bind("<Return>", lambda e: self._commit_total_gst_edit())
        self.total_gst_entry.bind(
            "<FocusOut>", lambda e: self.after_idle(self._commit_total_gst_edit)
        )
        tk.Label(gst, text="(auto / editable)", bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 8)).pack(side="left", padx=(4, 0))

        # Expandable cart
        self._tree_style()
        cols = ("S.No", "Product", "Qty", "Orig ₹", "Disc %", "Rate ₹", "Amount ₹")
        wrap = tk.Frame(parent, bg=PANEL)
        wrap.grid(row=1, column=0, sticky="nsew", padx=10, pady=(2, 4))
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        self.cart_tree = ttk.Treeview(
            wrap, columns=cols, show="headings", style="Angel.Treeview"
        )
        widths = {
            "S.No": 50, "Product": 260, "Qty": 70, "Orig ₹": 95,
            "Disc %": 70, "Rate ₹": 95, "Amount ₹": 110
        }
        for c in cols:
            self.cart_tree.heading(c, text=c)
            self.cart_tree.column(
                c, width=widths[c], minwidth=45,
                stretch=(c == "Product"),
                anchor="w" if c == "Product" else "center"
            )
        self.cart_tree.grid(row=0, column=0, sticky="nsew")
        sb = tk.Scrollbar(wrap, command=self.cart_tree.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.cart_tree.configure(yscrollcommand=sb.set)

        self.cart_tree.bind("<Double-1>", self._edit_cart_cell)
        self.cart_tree.bind("<Return>", self._edit_selected_qty)
        self.cart_tree.bind("<Delete>", lambda e: self._remove_cart_row())
        self.cart_tree.bind("<Escape>", lambda e: self._cancel_cart_editor())

        hint = tk.Label(
            parent,
            text="Double-click Product / Qty / Orig / Disc / Rate to edit • Amount automatic • Qty + Enter → Search",
            bg=PANEL, fg=MUTED, font=("Segoe UI", 9)
        )
        hint.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 3))

        # Fixed totals section.
        tot = tk.Frame(parent, bg=BG)
        tot.grid(row=3, column=0, sticky="ew", padx=10, pady=(0, 5))
        info = tk.Frame(tot, bg=BG)
        info.pack(side="left", fill="y")

        self.lbl_orig_total = self._tot_row(info, "Sub Total Value", "0.00")
        self.lbl_discount_total = self._tot_row(info, "Total Discount", "0.00")
        self.lbl_sub = self._tot_row(info, "Sub Total", "0.00")

        packrow = tk.Frame(info, bg=BG)
        packrow.pack(fill="x", pady=2)
        tk.Label(packrow, text="Packing Charges", bg=BG, fg=MUTED,
                 font=FONT, width=15, anchor="w").pack(side="left")
        self.packing_charge = tk.StringVar(value="0.00")
        self.packing_entry = tk.Entry(
            packrow, textvariable=self.packing_charge,
            bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
            relief="flat", font=FONT_B, width=12, justify="right"
        )
        self.packing_entry.pack(side="right", ipady=3)
        self.packing_entry.bind("<Return>", lambda e: self.refresh_cart())
        self.packing_entry.bind("<FocusOut>", lambda e: self.refresh_cart())

        self.lbl_cgst_total = self._tot_row(info, "CGST", "₹ 0.00")
        self.lbl_sgst_total = self._tot_row(info, "SGST", "₹ 0.00")
        self.lbl_gst_total = self._tot_row(info, "Total GST", "₹ 0.00")
        self.lbl_rnd = self._tot_row(info, "Round Off", "0.00")

        grand = tk.Frame(tot, bg=ACCENT, padx=18, pady=7)
        grand.pack(side="right")
        tk.Label(grand, text="GRAND TOTAL", bg=ACCENT, fg="#111",
                 font=FONT_B).pack(anchor="e")
        self.lbl_grand = tk.Label(
            grand, text="₹ 0.00", bg=ACCENT, fg="#111", font=FONT_BIG
        )
        self.lbl_grand.pack(anchor="e")

        # Fixed bottom action bar.
        btns = tk.Frame(parent, bg=PANEL)
        btns.grid(row=4, column=0, sticky="ew", padx=10, pady=(0, 7))
        tk.Button(
            btns, text="✚ Manual Item", bg=INFO, fg="white",
            font=FONT_B, relief="flat", padx=12, pady=8,
            command=self._add_manual_item
        ).pack(side="left", padx=3)
        tk.Button(
            btns, text="🗑 Remove", bg=DANGER, fg="white",
            font=FONT_B, relief="flat", padx=12, pady=8,
            command=self._remove_cart_row
        ).pack(side="left", padx=3)
        if self.editing_bill_no:
            tk.Button(
                btns, text="↩ Cancel Edit", bg="#64748b", fg="white",
                font=FONT_B, relief="flat", padx=12, pady=8,
                command=self._cancel_bill_edit
            ).pack(side="left", padx=3)
        tk.Button(
            btns, text="💾 Save + PDF (F9)", bg=SUCCESS, fg="white",
            font=FONT_B, relief="flat", padx=14, pady=8,
            command=self.save_bill
        ).pack(side="right", padx=3)

        self.refresh_cart()
    def _mark_total_gst_manual(self, event=None):
        """Any real keyboard edit makes Total GST a manual override."""
        if self.billing_user.get() == "User 2":
            self.total_gst_manual = True

    def _commit_total_gst_edit(self):
        """Validate Total GST. Blank means return to automatic calculation."""
        if not hasattr(self, "total_gst_override"):
            return "break"
        raw = self.total_gst_override.get().strip()
        if not raw:
            self.total_gst_manual = False
            self.refresh_cart()
            return "break"
        try:
            value = float(raw.replace(",", "").replace("₹", "").strip())
            if value < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid Total GST",
                                 "Total GST must be a non-negative number.")
            self.total_gst_override.set(
                f"{self._current_cart_totals().get('calculated_gst', 0.0):.2f}"
            )
            self.total_gst_manual = False
            self.refresh_cart()
            return "break"

        self.total_gst_override.set(f"{value:.2f}")
        self.total_gst_manual = self.billing_user.get() == "User 2"
        self.refresh_cart()
        return "break"

    def _current_packing_charge(self):
        try:
            return max(0.0, float(self.packing_charge.get() or 0))
        except Exception:
            return 0.0

    def _current_cart_totals(self):
        data = {
            "items": self.cart,
            "packing_charge": self._current_packing_charge(),
            "billing_user": getattr(self, "billing_user", tk.StringVar(value="User 1")).get(),
            "cgst_pct": getattr(self, "cgst_pct", tk.StringVar(value="9")).get(),
            "sgst_pct": getattr(self, "sgst_pct", tk.StringVar(value="9")).get(),
            "total_gst_manual": bool(getattr(self, "total_gst_manual", False)),
            "total_gst_override": getattr(
                self, "total_gst_override", tk.StringVar(value="0.00")
            ).get(),
        }
        return _bill_totals(data)

    def _on_billing_user_changed(self, event=None):
        """Apply the separate User 3 transport/credit Bill Of Supply mode."""
        if not hasattr(self, "billing_user"):
            return
        user = self.billing_user.get().strip()
        if user == "User 3":
            if not getattr(self, "editing_bill_no", None):
                if hasattr(self, "v_bill_no"):
                    self.v_bill_no.set(self._next_user3_bill_no())
                if hasattr(self, "bill_type"):
                    self.bill_type.set("Bill Of Supply")
            self._update_billing_user_mode()
        else:
            self._update_billing_user_mode()
            if not getattr(self, "editing_bill_no", None) and hasattr(self, "v_bill_no"):
                self.v_bill_no.set(self._next_bill_no(self.bill_type.get()))

    def _update_billing_user_mode(self):
        """User 1 is unchanged; User 2 alone displays and applies GST."""
        if not hasattr(self, "billing_user") or not hasattr(self, "gst_frame"):
            return
        if self.billing_user.get() == "User 2":
            # GST lives inside the fixed top billing section, so enabling it
            # never pushes the cart/totals/action bar out of view.
            self.gst_frame.pack(fill="x", pady=(0, 2))
        else:
            self.gst_frame.pack_forget()
            self.total_gst_manual = False
        self.refresh_cart()

    def _update_payment_from_amount(self):
        if not hasattr(self, "amount_paid"):
            return
        total = float(self._current_cart_totals()["grand_total"])
        try:
            paid = float(self.amount_paid.get() or 0)
        except (TypeError, ValueError):
            paid = 0.0
        paid = max(0.0, min(total, paid))
        self.amount_paid.set(f"{paid:.2f}")
        if total > 0 and paid >= total - 0.005:
            self.payment_status.set("Paid")
        elif paid <= 0.005:
            self.payment_status.set("Not Paid")
        else:
            self.payment_status.set("Partial")
        if hasattr(self, "payment_balance_label"):
            balance = max(0.0, total - paid)
            self.payment_balance_label.config(text=f"Balance: ₹ {balance:,.2f}")

    def _apply_payment_status(self):
        # Paid / Not Paid set their corresponding amounts. Partial is a
        # selectable state and keeps an editable amount below the grand total.
        status = self.payment_status.get().strip() if hasattr(self, "payment_status") else "Paid"
        total = float(self._current_cart_totals()["grand_total"])
        if status == "Paid":
            self.amount_paid.set(f"{total:.2f}")
        elif status == "Not Paid":
            self.amount_paid.set("0.00")
        else:
            try:
                current = float(self.amount_paid.get() or 0)
            except (TypeError, ValueError):
                current = 0.0
            if current <= 0 or (total > 0 and current >= total):
                self.amount_paid.set("0.00")
            # Keep Partial selected until the user enters a partial amount.
        try:
            shown_paid = float(self.amount_paid.get() or 0)
        except (TypeError, ValueError):
            shown_paid = 0.0
        if hasattr(self, "payment_balance_label"):
            self.payment_balance_label.config(
                text=f"Balance: ₹ {max(0.0, total - shown_paid):,.2f}"
            )

    def _cancel_bill_edit(self):
        self.editing_bill_no = None
        self.editing_bill_original = None
        self.show_billing()

    def _tot_row(self, parent, label, value):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=2)
        tk.Label(row, text=label, bg=BG, fg=MUTED, font=FONT,
                 width=12, anchor="w").pack(side="left")
        v = tk.Label(row, text=value, bg=BG, fg=TEXT, font=FONT_B,
                     width=12, anchor="e")
        v.pack(side="right")
        return v

    def _search_cart_rows(self, event=None):
        """Find an already-added bill item and bring it into view.

        The search is intentionally simple and forgiving: any typed text
        contained in the product name is treated as a match. The first match
        is selected; additional matches remain available by pressing Down /
        Up in the search box.
        """
        if not hasattr(self, "cart_tree") or not hasattr(self, "cart_search_var"):
            return

        query = self.cart_search_var.get().strip().lower()
        children = self.cart_tree.get_children()

        if not query:
            self.cart_tree.selection_remove(self.cart_tree.selection())
            return

        matches = []
        for iid in children:
            values = self.cart_tree.item(iid, "values")
            product_name = str(values[1] if len(values) > 1 else "").lower()
            if query in product_name:
                matches.append(iid)

        if matches:
            iid = matches[0]
            self.cart_tree.selection_set(iid)
            self.cart_tree.focus(iid)
            self.cart_tree.see(iid)
        else:
            self.cart_tree.selection_remove(self.cart_tree.selection())

    def _focus_searched_cart_row(self, event=None):
        """Enter from Cart Search focuses the matched cart row."""
        self._search_cart_rows()
        sel = self.cart_tree.selection()
        if sel:
            self.cart_tree.focus_set()
            self.cart_tree.see(sel[0])
        return "break"

    def _clear_cart_search(self, event=None):
        if hasattr(self, "cart_search_var"):
            self.cart_search_var.set("")
        if hasattr(self, "cart_tree"):
            self.cart_tree.selection_remove(self.cart_tree.selection())
        if hasattr(self, "cart_search_entry"):
            self.cart_search_entry.focus_set()
        return "break"

    def refresh_cart(self):
        if not hasattr(self, "cart_tree"):
            return
        self._cancel_cart_editor()
        for r in self.cart_tree.get_children():
            self.cart_tree.delete(r)

        for i, it in enumerate(self.cart, 1):
            qty = float(it.get("qty", 0) or 0)
            rate = float(it.get("rate", 0) or 0)
            amt = qty * rate
            self.cart_tree.insert(
                "", "end", iid=str(i - 1),
                values=(i, it.get("name", ""), f"{qty:g}",
                        f"{float(it.get('orig_rate', 0) or 0):.2f}",
                        f"{float(it.get('disc_pct', 0) or 0):g}%",
                        f"{rate:.2f}", f"{amt:.2f}")
            )

        totals = self._current_cart_totals()
        self.lbl_orig_total.config(text=f"₹ {totals['original_value']:,.2f}")
        self.lbl_discount_total.config(text=f"₹ {totals['discount_value']:,.2f}")
        self.lbl_sub.config(text=f"₹ {totals['sub_total']:,.2f}")
        user2 = totals.get("billing_user") == "User 2"
        self.lbl_cgst_total.config(text=f"₹ {totals['cgst']:,.2f}" if user2 else "₹ 0.00")
        self.lbl_sgst_total.config(text=f"₹ {totals['sgst']:,.2f}" if user2 else "₹ 0.00")
        self.lbl_gst_total.config(text=f"₹ {totals['total_gst']:,.2f}" if user2 else "₹ 0.00")
        if hasattr(self, "total_gst_override"):
            # Automatic until the user edits the Total GST entry.  Once edited,
            # keep the entered amount and use it in the PDF/grand total.
            if user2 and not self.total_gst_manual:
                self.total_gst_override.set(f"{totals['calculated_gst']:,.2f}")
            elif not user2:
                self.total_gst_override.set("0.00")
        self.lbl_rnd.config(text=f"₹ {totals['roundoff']:,.2f}")
        self.lbl_grand.config(text=f"₹ {totals['grand_total']:,.2f}")

        if hasattr(self, "payment_status"):
            self._apply_payment_status()

    def _edit_selected_qty(self, event=None):
        sel = self.cart_tree.selection()
        if not sel:
            return "break"
        self._start_qty_edit(int(sel[0]))
        return "break"

    def _edit_cart_cell(self, event=None):
        """Inline editor over cart cell. No Toplevel/new window."""
        if event is None:
            return "break"
        if self.cart_tree.identify("region", event.x, event.y) != "cell":
            return "break"

        row_id = self.cart_tree.identify_row(event.y)
        col_id = self.cart_tree.identify_column(event.x)
        if not row_id or not col_id:
            return "break"

        col_index = int(col_id[1:]) - 1
        # S.No and Amount are readonly.
        field = {
            1: "Product", 2: "Qty", 3: "Orig ₹",
            4: "Disc %", 5: "Rate ₹"
        }.get(col_index)
        if field:
            self._start_cart_edit(int(row_id), field)
        return "break"

    def _start_cart_edit(self, idx, field):
        if idx < 0 or idx >= len(self.cart):
            return
        self._cancel_cart_editor()

        col_map = {
            "Product": "#2", "Qty": "#3", "Orig ₹": "#4",
            "Disc %": "#5", "Rate ₹": "#6"
        }

        # Always bring the selected row into view before placing the inline
        # editor. Without this, rows below the visible area can receive an
        # editor that is effectively hidden, especially when adding item 17+
        # and immediately entering Qty.
        self.cart_tree.see(str(idx))
        self.update_idletasks()
        bbox = self.cart_tree.bbox(str(idx), col_map[field])
        if not bbox:
            return

        it = self.cart[idx]
        if field == "Product":
            value = str(it.get("name", ""))
        elif field == "Qty":
            value = f"{float(it.get('qty', 0) or 0):g}"
        elif field == "Orig ₹":
            value = f"{float(it.get('orig_rate', 0) or 0):.2f}"
        elif field == "Disc %":
            value = f"{float(it.get('disc_pct', 0) or 0):g}"
        else:
            value = f"{float(it.get('rate', 0) or 0):.2f}"

        x, y, w, h = bbox
        entry = tk.Entry(
            self.cart_tree, font=("Segoe UI", 10, "bold"),
            bg=ACCENT, fg="#111111", insertbackground="#111111",
            relief="solid", bd=2,
            justify="left" if field == "Product" else "right"
        )
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, value)
        entry.select_range(0, tk.END)
        entry.focus_set()

        self.cart_editor = entry
        self.cart_editor_item = idx
        self.cart_editor_col = field

        entry.bind(
            "<Return>",
            lambda e: self._commit_cart_edit(
                move_after=(field == "Qty"), editor=entry
            )
        )
        entry.bind(
            "<Tab>",
            lambda e: self._commit_cart_edit(
                tab_direction=1, editor=entry
            )
        )
        entry.bind(
            "<Shift-Key-Tab>",
            lambda e: self._commit_cart_edit(
                tab_direction=-1, editor=entry
            )
        )
        entry.bind("<Escape>", lambda e: self._cancel_cart_editor())
        entry.bind(
            "<FocusOut>",
            lambda e, ed=entry: self.after_idle(
                lambda: self._commit_cart_edit(editor=ed)
            )
        )

    def _commit_cart_edit(self, event=None, move_after=False,
                          tab_direction=0, editor=None):
        # Ignore a delayed FocusOut from an editor that has already been
        # replaced by the next Tab/Shift+Tab editor.
        if editor is not None and self.cart_editor is not editor:
            return "break"

        entry = self.cart_editor
        idx = self.cart_editor_item
        field = self.cart_editor_col
        if entry is None or idx is None or field is None:
            return "break"

        raw = entry.get().strip()
        current_idx = idx
        current_field = field
        self._cancel_cart_editor()
        if idx >= len(self.cart):
            return "break"

        it = self.cart[idx]

        if field == "Product":
            it["name"] = raw or "Manual Item"

        elif field == "Qty":
            try:
                qty = float(raw)
                if qty < 0:
                    raise ValueError
                it["qty"] = qty
            except ValueError:
                messagebox.showerror("Invalid Qty", "Qty must be a number.")
                self._start_qty_edit(idx)
                return "break"

        elif field == "Orig ₹":
            try:
                orig = float(raw)
                if orig < 0:
                    raise ValueError
                it["orig_rate"] = orig
                pct = 80.0 if it.get("manual") else float(it.get("disc_pct", 0) or 0)
                it["disc_pct"] = pct
                it["rate"] = round(orig * (1 - pct / 100.0), 2)
            except ValueError:
                messagebox.showerror(
                    "Invalid Rate", "Original Rate must be a number."
                )
                self._start_cart_edit(idx, "Orig ₹")
                return "break"

        elif field == "Disc %":
            try:
                pct = float(raw)
                if pct < 0 or pct > 100:
                    raise ValueError
                it["disc_pct"] = pct
                orig = float(it.get("orig_rate", 0) or 0)
                it["rate"] = round(orig * (1 - pct / 100.0), 2)
            except ValueError:
                messagebox.showerror(
                    "Invalid Discount",
                    "Discount % must be between 0 and 100."
                )
                self._start_cart_edit(idx, "Disc %")
                return "break"

        elif field == "Rate ₹":
            # Existing product items keep their original editable final-rate
            # behaviour. Manual items use the fixed 80% discount rule: typing
            # either Original Rate or Final Rate keeps both values consistent.
            try:
                rate = float(raw)
                if rate < 0:
                    raise ValueError
                if it.get("manual"):
                    it["rate"] = rate
                    it["disc_pct"] = 80.0
                    it["orig_rate"] = round(rate / 0.20, 2) if rate > 0 else 0.0
                else:
                    it["rate"] = rate
                    orig = float(it.get("orig_rate", 0) or 0)
                    it["disc_pct"] = (
                        round((1 - rate / orig) * 100, 2) if orig > 0 else 0
                    )
            except ValueError:
                messagebox.showerror(
                    "Invalid Rate", "Final Rate must be a number."
                )
                self._start_cart_edit(idx, "Rate ₹")
                return "break"

        self.refresh_cart()

        if tab_direction:
            # Tab/Shift+Tab walks through every editable cart field:
            # Product -> Qty -> Orig -> Disc -> Rate -> next/previous row.
            edit_fields = ["Product", "Qty", "Orig ₹", "Disc %", "Rate ₹"]
            try:
                pos = edit_fields.index(current_field)
            except ValueError:
                pos = 0

            next_pos = pos + tab_direction
            next_idx = current_idx

            if next_pos >= len(edit_fields):
                next_pos = 0
                next_idx += 1
            elif next_pos < 0:
                next_pos = len(edit_fields) - 1
                next_idx -= 1

            if 0 <= next_idx < len(self.cart):
                self.cart_tree.selection_set(str(next_idx))
                self.cart_tree.focus(str(next_idx))
                self._start_cart_edit(next_idx, edit_fields[next_pos])
            elif next_idx >= len(self.cart):
                # End of the cart: return to product search.
                self.search_var.set("")
                self.search_entry.focus_set()
                self.search_entry.icursor(tk.END)
            else:
                # Before the first cart field: return to product search.
                self.search_entry.focus_set()
                self.search_entry.icursor(tk.END)
        elif move_after:
            # Required workflow: Qty + Enter -> Search Product.
            self.search_var.set("")
            self.search_entry.focus_set()
            self.search_entry.icursor(tk.END)
        else:
            self.cart_tree.selection_set(str(idx))
            self.cart_tree.focus(str(idx))
        return "break"

    def _cancel_cart_editor(self, event=None):
        entry = getattr(self, "cart_editor", None)
        if entry is not None:
            try:
                entry.destroy()
            except Exception:
                pass
        self.cart_editor = None
        self.cart_editor_item = None
        self.cart_editor_col = None
        return "break"

    def _remove_cart_row(self):
        self._cancel_cart_editor()
        sel = self.cart_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if 0 <= idx < len(self.cart):
            del self.cart[idx]
        if hasattr(self, "cart_search_var"):
            self.cart_search_var.set("")
        self.refresh_cart()
        self.search_entry.focus_set()

    # ---------------------------------------------------------
    # Save bill
    # ---------------------------------------------------------
    def _validated_bill_date(self):
        """Return the editable bill date in DD-MM-YYYY format."""
        raw = str(getattr(self, "bill_date", tk.StringVar(value="")).get() or "").strip()
        if not raw:
            messagebox.showerror("Invalid Date", "Bill Date cannot be empty. Use DD-MM-YYYY.")
            if hasattr(self, "bill_date_entry"):
                self.bill_date_entry.focus_set()
            return None
        try:
            dt = datetime.datetime.strptime(raw, "%d-%m-%Y")
        except ValueError:
            messagebox.showerror(
                "Invalid Date",
                "Bill Date must be in DD-MM-YYYY format. Example: 04-10-2026."
            )
            if hasattr(self, "bill_date_entry"):
                self.bill_date_entry.focus_set()
            return None
        normalized = dt.strftime("%d-%m-%Y")
        self.bill_date.set(normalized)
        return normalized

    def save_bill(self):
        self._cancel_cart_editor()
        bill_date = self._validated_bill_date()
        if bill_date is None:
            return
        if not self.cart:
            messagebox.showwarning("Empty", "Add at least one product")
            return

        for i, it in enumerate(self.cart, 1):
            if not str(it.get("name", "")).strip():
                messagebox.showerror(
                    "Invalid Item", f"Product name is empty on row {i}."
                )
                self._start_cart_edit(i - 1, "Product")
                return
            try:
                qty = float(it.get("qty", 0))
                rate = float(it.get("rate", 0))
                orig = float(it.get("orig_rate", 0))
            except (TypeError, ValueError):
                messagebox.showerror(
                    "Invalid Item", f"Check Qty / Rate on row {i}."
                )
                return
            if qty < 0 or rate < 0 or orig < 0:
                messagebox.showerror(
                    "Invalid Item",
                    f"Negative values are not allowed on row {i}."
                )
                return

            # Manual items always use the 80% discount rule. Recalculate the
            # final rate from Original Rate so the saved amount is consistent.
            if it.get("manual"):
                it["disc_pct"] = 80.0
                it["rate"] = round(orig * 0.20, 2)

        bill_no = self.v_bill_no.get().strip()
        if not bill_no:
            messagebox.showerror("Invalid Bill Number", "Bill Number cannot be empty.")
            self.bill_no_entry.focus_set()
            return

        # Do not allow two different saved bills to use the same editable number.
        duplicate = next(
            (b for b in BILLS
             if str(b.get("bill_no", "")).strip() == bill_no
             and str(b.get("bill_no", "")).strip() != str(self.editing_bill_no or "").strip()),
            None
        )
        if duplicate is not None:
            messagebox.showerror(
                "Duplicate Bill Number",
                f"Bill Number {bill_no} already exists. Please enter a different number."
            )
            self.bill_no_entry.focus_set()
            return

        bill_type = self.bill_type.get().strip() if hasattr(self, "bill_type") else "Bill Of Supply"
        if bill_type not in ("Bill Of Supply", "Tax Invoice", "Estimate"):
            bill_type = "Bill Of Supply"

        packing = self._current_packing_charge()
        billing_user = self.billing_user.get().strip() if hasattr(self, "billing_user") else "User 1"
        if billing_user not in ("User 1", "User 2", "User 3", "User 4"):
            billing_user = "User 1"
        # User 3 is reserved for transport-purpose credit Bill Of Supply documents.
        if billing_user == "User 3":
            bill_type = "Bill Of Supply"
            if not self.editing_bill_no:
                bill_no = self._next_user3_bill_no()
                self.v_bill_no.set(bill_no)
        cgst_pct = _safe_number(self.cgst_pct.get(), 9) if billing_user == "User 2" else 0.0
        sgst_pct = _safe_number(self.sgst_pct.get(), 9) if billing_user == "User 2" else 0.0
        total_gst_manual = bool(getattr(self, "total_gst_manual", False)) and billing_user == "User 2"
        total_gst_override = _safe_number(
            getattr(self, "total_gst_override", tk.StringVar(value="0.00")).get(),
            0.0
        )
        calc = _bill_totals({
            "items": self.cart, "packing_charge": packing,
            "billing_user": billing_user, "cgst_pct": cgst_pct, "sgst_pct": sgst_pct,
            "total_gst_manual": total_gst_manual,
            "total_gst_override": total_gst_override,
        })
        sub = calc["sub_total"]
        gt = calc["grand_total"]
        rnd = calc["roundoff"]

        selected_status = (
            self.payment_status.get().strip()
            if hasattr(self, "payment_status") else "Paid"
        )
        if selected_status not in ("Paid", "Partial", "Not Paid"):
            selected_status = "Paid"

        try:
            paid = float(self.amount_paid.get() or 0)
        except (TypeError, ValueError):
            messagebox.showerror("Invalid Amount", "Amount Paid must be a number.")
            self.amount_paid_entry.focus_set()
            return

        if paid < 0:
            messagebox.showerror("Invalid Amount", "Amount Paid cannot be negative.")
            return

        if selected_status == "Paid":
            paid = float(gt)
            status = "Paid"
        elif selected_status == "Not Paid":
            paid = 0.0
            status = "Not Paid"
        else:
            if gt <= 0:
                messagebox.showerror(
                    "Invalid Payment",
                    "Partial payment is not possible when the bill total is zero."
                )
                return
            if paid <= 0.005 or paid >= gt - 0.005:
                messagebox.showerror(
                    "Invalid Partial Payment",
                    "For Partial payment, Amount Paid must be greater than 0 "
                    "and less than the Grand Total."
                )
                self.amount_paid_entry.focus_set()
                return
            status = "Partial"

        self.amount_paid.set(f"{paid:.2f}")
        self.payment_status.set(status)

        bill = {
            "bill_no": bill_no,
            "bill_type": bill_type,
            "date": bill_date,
            "customer": self.cust_name.get().strip(),
            "mobile": self.cust_mob.get().strip(),
            "address": self.cust_address.get("1.0", "end-1c").strip(),
            "gstin": (self.editing_bill_original or {}).get("gstin", ""),
            "items": copy.deepcopy(self.cart),
            "original_value": calc["original_value"],
            "discount": calc["discount_value"],
            "subtotal": calc["sub_total"],
            "packing_charge": calc["packing_charge"],
            "billing_user": billing_user,
            "cgst_pct": cgst_pct,
            "sgst_pct": sgst_pct,
            "cgst": calc["cgst"],
            "sgst": calc["sgst"],
            "total_gst": calc["total_gst"],
            "total_gst_manual": total_gst_manual,
            "total_gst_override": calc["total_gst"] if total_gst_manual else calc["calculated_gst"],
            "taxable_value": calc["taxable_value"],
            "roundoff": calc["roundoff"],
            "grand_total": calc["grand_total"],
            "payment_status": status,
            "amount_paid": paid,
        }

        # Every newly saved bill gets a unique timestamped PDF filename.
        # The customer name is also included so PDFs remain easy to identify.
        bill["pdf_timestamp"] = datetime.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        out_pdf = self._bill_pdf_path(bill, new_file=True)
        bill["pdf_path"] = out_pdf
        try:
            generate_a4_invoice(bill, out_pdf)
        except Exception as e:
            messagebox.showerror("PDF Error", str(e))
            return

        if self.editing_bill_no:
            old_bill = self.editing_bill_original or {}
            found = next((b for b in BILLS if b.get("bill_no") == self.editing_bill_no), None)
            if found is None:
                messagebox.showerror("Update Error", f"Bill {self.editing_bill_no} was not found.")
                return
            # Same bill number, same PDF path: only the existing record is replaced.
            found.clear()
            found.update(copy.deepcopy(bill))
            self._adjust_stock_for_bill_change(old_bill, bill)
            save_json(BILLS_FILE, BILLS)

            # Remove the previous PDF after an edit because this save creates
            # a fresh timestamped PDF. This also handles bill-number/type changes.
            old_pdf = str(old_bill.get("pdf_path", "") or "")
            if not old_pdf:
                old_pdf = self._history_bill_pdf(old_bill)
            if old_pdf != out_pdf and os.path.exists(old_pdf):
                try:
                    os.remove(old_pdf)
                except Exception:
                    pass

            messagebox.showinfo(
                "Updated",
                f"Bill {bill['bill_no']} updated successfully.\n\n"
                f"Total: ₹{gt:,.2f}\nPDF: {out_pdf}"
            )
            self.editing_bill_no = None
            self.editing_bill_original = None
            self.show_billing()
        else:
            BILLS.append(bill)
            save_json(BILLS_FILE, BILLS)
            self._update_stock_from_bill(bill)

            prefix = self._bill_type_prefix(bill.get("bill_type"))
            counters = SETTINGS.setdefault("bill_counters", {})
            try:
                current_num = int(str(bill["bill_no"]).rsplit("-", 1)[-1])
            except (ValueError, TypeError):
                current_num = 0
            counters[prefix] = max(
                int(counters.get(prefix, 1) or 1),
                current_num + 1
            )
            # Keep the old setting for compatibility with any unrelated code.
            SETTINGS["next_bill"] = SETTINGS.get("next_bill", 1) + 1
            save_json(SETTINGS_FILE, SETTINGS)

            messagebox.showinfo(
                "Saved",
                f"Bill saved!\n\nNo: {bill['bill_no']}\n"
                f"Total: ₹{gt:,.2f}\n\nPDF: {out_pdf}"
            )

            try:
                if os.name == "nt":
                    os.startfile(out_pdf)
                else:
                    os.system(f'xdg-open "{out_pdf}"')
            except Exception:
                pass

            self.cart = []
            self.cust_name.set("")
            self.cust_mob.set("")
            self.cust_address.delete("1.0", tk.END)
            self.packing_charge.set("0.00")
            self.billing_user.set("User 1")
            self.cgst_pct.set("9")
            self.sgst_pct.set("9")
            self.total_gst_manual = False
            self.total_gst_override.set("0.00")
            self.bill_type.set("Bill Of Supply")
            self._update_billing_user_mode()
            self.payment_status.set("Paid")
            self.amount_paid.set("0.00")
            self._apply_payment_status()
            self.v_bill_no.set(self._next_bill_no("Bill Of Supply"))
            self.refresh_cart()
            self.search_var.set("")
            self.search_entry.focus_set()

    # =========================================================
    # STATEMENT
    # =========================================================
    def _statement_pdf_path(self, period_label="Statement"):
        safe = re.sub(r"[^A-Za-z0-9_-]+", "_", str(period_label)).strip("_") or "Statement"
        stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(STATEMENT_DIR, f"Statement_{safe}_{stamp}.pdf")

    def _statement_generate_pdf(self, bills, title, start=None, end=None):
        """Generate an A4 sales statement for the supplied bills."""
        out_pdf = self._statement_pdf_path(title)
        c = canvas.Canvas(out_pdf, pagesize=A4)
        W, H = A4

        logo_path = LOGO_PATH if os.path.exists(LOGO_PATH) else LOGO_FALLBACK_PATH
        if os.path.exists(logo_path):
            try:
                c.drawImage(
                    logo_path, 15*mm, H-34*mm, width=35*mm, height=20*mm,
                    preserveAspectRatio=True, anchor="sw", mask="auto"
                )
            except Exception:
                pass

        if os.path.exists(ROUND_LOGO_PATH):
            try:
                c.drawImage(
                    ROUND_LOGO_PATH, W-42*mm, H-35*mm, width=26*mm, height=26*mm,
                    preserveAspectRatio=True, anchor="sw", mask="auto"
                )
            except Exception:
                pass

        c.setFillColor(colors.HexColor("#8B0000"))
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(W/2, H-14*mm, SETTINGS.get("shop_name", "ANGEL PATTASU KADAI"))
        c.setFillColor(colors.HexColor("#B8860B"))
        c.setFont("Helvetica-Oblique", 9)
        c.drawCentredString(W/2, H-19*mm, SETTINGS.get("tagline", "Most Reliable Brand"))
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 8.2)
        c.drawCentredString(W/2, H-24*mm, SETTINGS.get("address", ""))
        c.drawCentredString(
            W/2, H-28.5*mm,
            f"Phone: {SETTINGS.get('phone','')}   |   Email: {SETTINGS.get('email','')}"
        )
        c.drawCentredString(
            W/2, H-33*mm,
            f"GSTIN: {SETTINGS.get('gstin','')}   |   State: {SETTINGS.get('state','')}"
        )

        c.setStrokeColor(colors.HexColor("#8B0000"))
        c.setLineWidth(1)
        c.line(15*mm, H-36.5*mm, W-15*mm, H-36.5*mm)

        c.setFillColor(colors.HexColor("#8B0000"))
        c.setFont("Helvetica-Bold", 14)
        c.drawCentredString(W/2, H-45*mm, "SALES STATEMENT")

        period_text = title
        if start and end:
            period_text = f"{title}  |  {start.strftime('%d-%m-%Y')} to {end.strftime('%d-%m-%Y')}"
        elif start:
            period_text = f"{title}  |  From {start.strftime('%d-%m-%Y')}"
        elif end:
            period_text = f"{title}  |  Up to {end.strftime('%d-%m-%Y')}"

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 9)
        c.drawCentredString(W/2, H-51*mm, period_text)
        c.setFont("Helvetica", 7.5)
        c.drawRightString(W-15*mm, H-56*mm,
                           f"Generated: {datetime.datetime.now().strftime('%d-%m-%Y %I:%M %p')}")

        # Statement columns: Date | Bill No | Customer | Mobile | Amount | Paid | Balance
        left = 12*mm
        right = W-12*mm
        # Keep the total column width inside the printable A4 width.
        # Previous widths exceeded the available area, which clipped the
        # Balance column on the right side of the PDF.
        col_w = [20*mm, 26*mm, 43*mm, 29*mm, 22*mm, 22*mm, 22*mm]
        labels = ["Date", "Bill No", "Customer", "Mobile", "Amount", "Paid", "Balance"]
        header_h = 7*mm
        top = H-62*mm
        bottom = top-header_h

        def draw_header(y):
            c.setFillColor(colors.HexColor("#8B0000"))
            c.rect(left, y-header_h, right-left, header_h, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 7.2)
            x = left
            for w, label in zip(col_w, labels):
                if label == "Customer":
                    c.drawString(x+1.5*mm, y-header_h+2.2*mm, label)
                else:
                    c.drawCentredString(x+w/2, y-header_h+2.2*mm, label)
                x += w
            return y-header_h

        ry = draw_header(top)
        row_h = 6.5*mm
        page_count = 0
        total_sales = 0.0
        total_paid = 0.0
        total_balance = 0.0

        for b in bills:
            status, paid, balance = self._bill_payment_values(b)
            total = float(_bill_totals(b)["grand_total"])
            total_sales += total
            total_paid += paid
            total_balance += balance

            if ry-row_h < 34*mm:
                c.setStrokeColor(colors.HexColor("#cccccc"))
                c.line(left, ry, right, ry)
                c.showPage()
                page_count += 1
                c.setFillColor(colors.HexColor("#8B0000"))
                c.setFont("Helvetica-Bold", 11)
                c.drawCentredString(W/2, H-14*mm, "SALES STATEMENT - CONTINUED")
                c.setFillColor(colors.black)
                c.setFont("Helvetica", 7.5)
                c.drawCentredString(W/2, H-18.5*mm, period_text)
                ry = draw_header(H-24*mm)

            c.setStrokeColor(colors.HexColor("#dddddd"))
            c.setLineWidth(0.25)
            c.line(left, ry-row_h, right, ry-row_h)
            vals = [
                str(b.get("date", "-") or "-"),
                str(b.get("bill_no", "-") or "-"),
                str(b.get("customer", "-") or "-"),
                str(b.get("mobile", "-") or "-"),
                f"Rs. {total:,.2f}",
                f"Rs. {paid:,.2f}",
                f"Rs. {balance:,.2f}",
            ]

            c.setFillColor(colors.black)
            c.setFont("Helvetica", 7)
            x = left
            for w, label, val in zip(col_w, labels, vals):
                if label == "Customer":
                    customer = val
                    max_w = w - 3*mm
                    if stringWidth(customer, "Helvetica", 7) > max_w:
                        while customer and stringWidth(customer + "...", "Helvetica", 7) > max_w:
                            customer = customer[:-1]
                        customer = customer.rstrip() + "..."
                    c.drawString(x+1.5*mm, ry-row_h+2.15*mm, customer)
                else:
                    c.drawCentredString(x+w/2, ry-row_h+2.15*mm, val)
                x += w
            ry -= row_h

        c.setStrokeColor(colors.black)
        c.setLineWidth(0.6)
        c.line(left, ry, right, ry)

        # Summary section
        sy = ry - 7*mm
        c.setFillColor(colors.HexColor("#8B0000"))
        c.setFont("Helvetica-Bold", 10)
        c.drawString(left, sy, "STATEMENT SUMMARY")

        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 8.5)
        summary_rows = [
            ("Total Bills", len(bills)),
            ("Total Sales", f"Rs. {total_sales:,.2f}"),
            ("Total Paid", f"Rs. {total_paid:,.2f}"),
            ("Total Pending", f"Rs. {total_balance:,.2f}"),
        ]
        sy -= 6*mm
        for label, value in summary_rows:
            c.drawString(left+3*mm, sy, str(label))
            c.drawRightString(right-3*mm, sy, str(value))
            sy -= 4.5*mm

        c.setFillColor(colors.HexColor("#B8860B"))
        c.setFont("Helvetica-Bold", 8)
        # Use a complete number-to-words value for the statement.
        # The existing global helper returns only "Rupees Only" for values
        # below 100 because it expects the caller to supply the lower part.
        def statement_amount_words(value):
            try:
                n = int(round(float(value)))
            except (TypeError, ValueError):
                n = 0

            if n == 0:
                return "Zero Rupees Only"

            ones = [
                "", "One", "Two", "Three", "Four", "Five", "Six", "Seven",
                "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen",
                "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen",
                "Nineteen"
            ]
            tens = [
                "", "", "Twenty", "Thirty", "Forty", "Fifty",
                "Sixty", "Seventy", "Eighty", "Ninety"
            ]

            def two_words(x):
                if x < 20:
                    return ones[x]
                return tens[x // 10] + ((" " + ones[x % 10]) if x % 10 else "")

            def under_thousand(x):
                if x < 100:
                    return two_words(x)
                return (
                    ones[x // 100] + " Hundred" +
                    ((" " + two_words(x % 100)) if x % 100 else "")
                )

            parts = []
            crore, n2 = divmod(n, 10000000)
            lakh, n2 = divmod(n2, 100000)
            thousand, n2 = divmod(n2, 1000)
            hundred = n2

            if crore:
                parts.append(under_thousand(crore) + " Crore")
            if lakh:
                parts.append(under_thousand(lakh) + " Lakh")
            if thousand:
                parts.append(under_thousand(thousand) + " Thousand")
            if hundred:
                parts.append(under_thousand(hundred))

            return " ".join(parts) + " Rupees Only"

        c.drawString(
            left, sy-2*mm,
            f"Sales in words: {statement_amount_words(total_sales)}"
        )

        c.setStrokeColor(colors.HexColor("#8B0000"))
        c.setLineWidth(0.6)
        c.line(15*mm, 19*mm, W-15*mm, 19*mm)
        c.setFillColor(colors.HexColor("#8B0000"))
        c.setFont("Helvetica-Bold", 8.5)
        c.drawCentredString(W/2, 13.5*mm, "Thank You! Visit Again")
        c.setFillColor(colors.HexColor("#666666"))
        c.setFont("Helvetica", 6.5)
        c.drawCentredString(
            W/2, 9*mm,
            f"{SETTINGS.get('shop_name','')}  |  {SETTINGS.get('phone','')}  |  {SETTINGS.get('email','')}"
        )

        c.showPage()
        c.save()
        return out_pdf

    def show_statement(self):
        self.current_view = "show_statement"
        self.clear_content()

        head = tk.Frame(self.content, bg=BG)
        head.pack(fill="x", padx=20, pady=(18, 8))
        tk.Label(head, text="📊  Sales Statement", bg=BG, fg=TEXT,
                 font=FONT_H1).pack(side="left")
        tk.Label(head, text="Date-wise sales statement • Generate PDF",
                 bg=BG, fg=MUTED, font=FONT).pack(side="right")

        filters = tk.Frame(self.content, bg=PANEL)
        filters.pack(fill="x", padx=20, pady=(0, 8))

        tk.Label(filters, text="Period:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(12, 5), pady=12)

        self.statement_period = tk.StringVar(value="Today")
        values = [
            "Today", "Yesterday", "This Month", "Last Month",
            "This Year", "Last Year", "Custom Range", "All"
        ]
        self.statement_period_combo = ttk.Combobox(
            filters, textvariable=self.statement_period,
            values=values, state="readonly", width=16,
            font=("Segoe UI", 10)
        )
        self.statement_period_combo.pack(side="left", padx=5, pady=10, ipady=4)
        self.statement_period_combo.bind(
            "<<ComboboxSelected>>",
            lambda e: self._statement_period_changed()
        )

        tk.Label(filters, text="From:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(14, 4))
        self.statement_from = tk.StringVar()
        tk.Entry(filters, textvariable=self.statement_from, width=12,
                 bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
                 relief="flat", font=FONT).pack(side="left", padx=3, ipady=5)

        tk.Label(filters, text="To:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(8, 4))
        self.statement_to = tk.StringVar()
        tk.Entry(filters, textvariable=self.statement_to, width=12,
                 bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
                 relief="flat", font=FONT).pack(side="left", padx=3, ipady=5)

        tk.Label(filters, text="Search:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(14, 4))
        self.statement_search = tk.StringVar()
        self.statement_search.trace_add("write", lambda *_: self.refresh_statement())
        tk.Entry(filters, textvariable=self.statement_search, width=24,
                 bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
                 relief="flat", font=FONT).pack(side="left", padx=3, ipady=5)

        tk.Button(filters, text="🔎 Apply", bg=INFO, fg="white",
                  font=FONT_B, relief="flat", padx=12, pady=6,
                  command=self.refresh_statement).pack(side="left", padx=5)
        tk.Button(filters, text="✖ Clear", bg="#64748b", fg="white",
                  font=FONT_B, relief="flat", padx=10, pady=6,
                  command=self._clear_statement_filters).pack(side="left", padx=2)
        tk.Button(filters, text="📄 Generate PDF", bg=SUCCESS, fg="white",
                  font=FONT_B, relief="flat", padx=14, pady=7,
                  command=self.generate_statement_pdf).pack(side="right", padx=8)

        self.statement_summary = tk.Label(
            self.content, text="", bg=BG, fg=ACCENT, font=FONT_B, anchor="w"
        )
        self.statement_summary.pack(fill="x", padx=22, pady=(0, 6))

        self._tree_style()
        cols = ("Date", "Bill No", "Customer", "Mobile", "Amount", "Paid", "Balance")
        wrap = tk.Frame(self.content, bg=PANEL)
        wrap.pack(fill="both", expand=True, padx=20, pady=(0, 20))

        self.statement_tree = ttk.Treeview(
            wrap, columns=cols, show="headings", style="Angel.Treeview"
        )
        widths = {
            "Date": 120, "Bill No": 130, "Customer": 260,
            "Mobile": 150, "Amount": 130, "Paid": 130, "Balance": 130
        }
        for c in cols:
            self.statement_tree.heading(c, text=c)
            self.statement_tree.column(
                c, width=widths[c],
                anchor="w" if c == "Customer" else "center"
            )
        self.statement_tree.tag_configure(
            "unpaid", background="#facc15", foreground="#111111"
        )
        self.statement_tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        sb = tk.Scrollbar(wrap, command=self.statement_tree.yview)
        self.statement_tree.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        self._statement_period_changed()

    def _statement_parse_date(self, value):
        return self._history_parse_date(value)

    def _statement_period_changed(self):
        period = self.statement_period.get()
        today = datetime.date.today()

        if period == "Today":
            start = end = today
        elif period == "Yesterday":
            d = today - datetime.timedelta(days=1)
            start = end = d
        elif period == "This Month":
            start = today.replace(day=1)
            end = today
        elif period == "Last Month":
            first_this = today.replace(day=1)
            end = first_this - datetime.timedelta(days=1)
            start = end.replace(day=1)
        elif period == "This Year":
            start = datetime.date(today.year, 1, 1)
            end = today
        elif period == "Last Year":
            start = datetime.date(today.year - 1, 1, 1)
            end = datetime.date(today.year - 1, 12, 31)
        elif period == "All":
            start = end = None
        else:
            # Custom Range: keep the user's manually entered values.
            self.refresh_statement()
            return

        self.statement_from.set(start.strftime("%d-%m-%Y") if start else "")
        self.statement_to.set(end.strftime("%d-%m-%Y") if end else "")
        self.refresh_statement()

    def _clear_statement_filters(self):
        self.statement_period.set("Today")
        self.statement_search.set("")
        self._statement_period_changed()

    def _statement_date_range(self):
        period = self.statement_period.get()
        if period == "All":
            return None, None

        start = self._statement_parse_date(self.statement_from.get())
        end = self._statement_parse_date(self.statement_to.get())

        if period == "Custom Range":
            return start, end
        return start, end

    def _statement_filtered_bills(self):
        start, end = self._statement_date_range()
        q = self.statement_search.get().lower().strip()
        shown = []

        for b in BILLS:
            bd = self._statement_parse_date(b.get("date"))
            if start and (bd is None or bd < start):
                continue
            if end and (bd is None or bd > end):
                continue

            hay = (
                f"{b.get('bill_no','')} {b.get('customer','')} "
                f"{b.get('mobile','')} {b.get('address','')} "
                f"{b.get('date','')} {b.get('payment_status','')} "
                f"{b.get('grand_total','')}"
            ).lower()
            if q and q not in hay:
                continue
            shown.append(b)

        # Newest first, matching Bill History behaviour.
        shown.sort(
            key=lambda b: (
                self._statement_parse_date(b.get("date")) or datetime.date.min,
                str(b.get("bill_no", ""))
            ),
            reverse=True
        )
        return shown

    def refresh_statement(self):
        if not hasattr(self, "statement_tree"):
            return

        for r in self.statement_tree.get_children():
            self.statement_tree.delete(r)

        shown = self._statement_filtered_bills()
        total_sales = total_paid = total_balance = 0.0

        for b in shown:
            status, paid, balance = self._bill_payment_values(b)
            total = float(_bill_totals(b)["grand_total"])
            total_sales += total
            total_paid += paid
            total_balance += balance
            tag = "unpaid" if balance > 0 else "paid"

            self.statement_tree.insert(
                "", "end", tags=(tag,),
                values=(
                    b.get("date", "-") or "-",
                    b.get("bill_no", "-") or "-",
                    b.get("customer", "-") or "-",
                    b.get("mobile", "-") or "-",
                    f"₹ {total:,.2f}",
                    f"₹ {paid:,.2f}",
                    f"₹ {balance:,.2f}",
                )
            )

        if hasattr(self, "statement_summary"):
            self.statement_summary.config(
                text=(
                    f"Showing {len(shown)} bills   |   "
                    f"Sales: ₹ {total_sales:,.2f}   |   "
                    f"Paid: ₹ {total_paid:,.2f}   |   "
                    f"Pending: ₹ {total_balance:,.2f}"
                )
            )

    def generate_statement_pdf(self):
        period = self.statement_period.get()
        start, end = self._statement_date_range()

        if period == "Custom Range" and not start and not end:
            messagebox.showwarning(
                "Invalid Date",
                "Enter a valid From and To date.\n\n"
                "Example: 01-10-2026"
            )
            return
        if start and end and start > end:
            messagebox.showwarning(
                "Invalid Date Range",
                "From date cannot be later than To date."
            )
            return

        bills = self._statement_filtered_bills()
        if period == "All":
            title = "All Sales"
        elif period == "Custom Range":
            title = "Custom Range"
        else:
            title = period

        try:
            out_pdf = self._statement_generate_pdf(
                bills, title, start=start, end=end
            )
            if not self._open_pdf_file(out_pdf):
                messagebox.showinfo(
                    "Statement PDF Generated",
                    f"Statement PDF generated successfully.\n\n{out_pdf}"
                )
        except Exception as e:
            messagebox.showerror("Statement PDF Error", str(e))

    # =========================================================
    # HISTORY
    # =========================================================
    def show_history(self):
        self.current_view = "show_history"
        self.clear_content()

        head = tk.Frame(self.content, bg=BG)
        head.pack(fill="x", padx=20, pady=(18, 8))
        tk.Label(head, text="📜  Bill History", bg=BG, fg=TEXT,
                 font=FONT_H1).pack(side="left")
        tk.Label(head, text="Double-click = Edit in New Bill • Unpaid = Yellow",
                 bg=BG, fg=MUTED, font=FONT).pack(side="right")

        # ---- Filters ----
        filters = tk.Frame(self.content, bg=PANEL)
        filters.pack(fill="x", padx=20, pady=(0, 8))

        tk.Label(filters, text="Period:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(12, 5), pady=12)
        self.hist_period = tk.StringVar(value="All")
        period_values = [
            "All", "Today", "Yesterday", "This Month", "Last Month",
            "2025", "2026", "Custom Range"
        ]
        self.hist_period_combo = ttk.Combobox(
            filters, textvariable=self.hist_period,
            values=period_values, state="readonly", width=15,
            font=("Segoe UI", 10)
        )
        self.hist_period_combo.pack(side="left", padx=5, pady=10, ipady=4)
        self.hist_period_combo.bind("<<ComboboxSelected>>", lambda e: self._history_period_changed())

        # User filter: quickly view bills created by a specific billing user.
        tk.Label(filters, text="User:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(14, 4), pady=12)
        self.hist_user = tk.StringVar(value="All Users")
        self.hist_user_combo = ttk.Combobox(
            filters, textvariable=self.hist_user,
            values=["All Users", "User 1", "User 2", "User 3", "User 4"],
            state="readonly", width=12, font=("Segoe UI", 10)
        )
        self.hist_user_combo.pack(side="left", padx=4, pady=10, ipady=4)
        self.hist_user_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_history())

        # Bill type filter: quickly view Bill Of Supply / Tax Invoice / Estimate.
        tk.Label(filters, text="Bill Type:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(10, 4), pady=12)
        self.hist_bill_type = tk.StringVar(value="All Types")
        self.hist_bill_type_combo = ttk.Combobox(
            filters, textvariable=self.hist_bill_type,
            values=["All Types", "Bill Of Supply", "Tax Invoice", "Estimate"],
            state="readonly", width=16, font=("Segoe UI", 10)
        )
        self.hist_bill_type_combo.pack(side="left", padx=4, pady=10, ipady=4)
        self.hist_bill_type_combo.bind("<<ComboboxSelected>>", lambda e: self.refresh_history())

        tk.Label(filters, text="From:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(14, 4))
        self.hist_from = tk.StringVar(value="")
        tk.Entry(filters, textvariable=self.hist_from, width=12,
                 bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
                 relief="flat", font=FONT).pack(side="left", padx=3, ipady=5)
        tk.Label(filters, text="To:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(8, 4))
        self.hist_to = tk.StringVar(value="")
        tk.Entry(filters, textvariable=self.hist_to, width=12,
                 bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
                 relief="flat", font=FONT).pack(side="left", padx=3, ipady=5)

        tk.Label(filters, text="Search:", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(16, 4))
        self.hist_search = tk.StringVar()
        self.hist_search.trace_add("write", lambda *_: self.refresh_history())
        tk.Entry(filters, textvariable=self.hist_search, width=28,
                 bg=INPUT_BG, fg=TEXT, insertbackground=TEXT,
                 relief="flat", font=FONT).pack(side="left", padx=3, ipady=5)

        tk.Button(filters, text="🔎 Apply", bg=INFO, fg="white",
                  font=FONT_B, relief="flat", padx=12, pady=6,
                  command=self.refresh_history).pack(side="left", padx=6)
        tk.Button(filters, text="✖ Clear", bg="#64748b", fg="white",
                  font=FONT_B, relief="flat", padx=10, pady=6,
                  command=self._clear_history_filters).pack(side="left", padx=2)

        self.hist_summary = tk.Label(
            self.content, text="", bg=BG, fg=ACCENT, font=FONT_B, anchor="w"
        )
        self.hist_summary.pack(fill="x", padx=22, pady=(0, 6))

        self._tree_style()
        self.hist_cols = ("Bill No", "Date", "Customer", "Mobile", "Items",
                          "Total", "Payment", "Paid", "Balance")
        wrap = tk.Frame(self.content, bg=PANEL)
        wrap.pack(fill="both", expand=True, padx=20, pady=(0, 10))
        self.hist_tree = ttk.Treeview(
            wrap, columns=self.hist_cols, show="headings", style="Angel.Treeview"
        )
        widths = {
            "Bill No": 110, "Date": 105, "Customer": 190, "Mobile": 125,
            "Items": 60, "Total": 105, "Payment": 95, "Paid": 100, "Balance": 105
        }
        for c in self.hist_cols:
            self.hist_tree.heading(c, text=c)
            self.hist_tree.column(c, width=widths[c],
                                  anchor="w" if c == "Customer" else "center")
        # Pending / Not Paid rows use pink instead of the previous yellow.
        self.hist_tree.tag_configure("unpaid", background="#f9a8d4", foreground="#111111")
        self.hist_tree.tag_configure("paid", background=INPUT_BG, foreground=TEXT)
        self.hist_tree.pack(side="left", fill="both", expand=True, padx=10, pady=10)
        sb = tk.Scrollbar(wrap, command=self.hist_tree.yview)
        self.hist_tree.config(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")

        self.hist_tree.bind("<Double-1>", self._history_double_click)
        self.hist_tree.bind("<Button-3>", self._history_right_click)
        self.hist_tree.bind("<Return>", lambda e: self.edit_history_bill())

        btns = tk.Frame(self.content, bg=BG)
        btns.pack(fill="x", padx=20, pady=(0, 20))
        tk.Button(btns, text="✏ Edit in New Bill", bg=INFO, fg="white",
                  font=FONT_B, relief="flat", padx=14, pady=8,
                  command=self.edit_history_bill).pack(side="left", padx=4)
        tk.Button(btns, text="🖨 Print", bg=ACCENT, fg="#111",
                  font=FONT_B, relief="flat", padx=14, pady=8,
                  command=self.print_history_bill).pack(side="left", padx=4)
        tk.Button(btns, text="📄 Open PDF", bg="#64748b", fg="white",
                  font=FONT_B, relief="flat", padx=14, pady=8,
                  command=self.open_history_pdf).pack(side="left", padx=4)
        tk.Button(btns, text="🗑 Delete Bill", bg=DANGER, fg="white",
                  font=FONT_B, relief="flat", padx=14, pady=8,
                  command=self.delete_bill).pack(side="left", padx=4)
        tk.Button(btns, text="🔄 Refresh", bg=INFO, fg="white",
                  font=FONT_B, relief="flat", padx=14, pady=8,
                  command=self.refresh_history).pack(side="right", padx=4)

        self.refresh_history()

    def _history_parse_date(self, value):
        s = str(value or "").strip()
        if not s:
            return None
        for fmt in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d",
                    "%d-%m-%Y %H:%M:%S", "%d/%m/%Y %H:%M:%S"):
            try:
                return datetime.datetime.strptime(s, fmt).date()
            except ValueError:
                pass
        m = re.search(r"(\\d{1,2})[-/](\\d{1,2})[-/](\\d{4})", s)
        if m:
            try:
                return datetime.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                pass
        return None

    def _history_period_changed(self):
        period = self.hist_period.get()
        today = datetime.date.today()
        if period == "Today":
            self.hist_from.set(today.strftime("%d-%m-%Y"))
            self.hist_to.set(today.strftime("%d-%m-%Y"))
        elif period == "Yesterday":
            d = today - datetime.timedelta(days=1)
            self.hist_from.set(d.strftime("%d-%m-%Y"))
            self.hist_to.set(d.strftime("%d-%m-%Y"))
        elif period == "This Month":
            first = today.replace(day=1)
            self.hist_from.set(first.strftime("%d-%m-%Y"))
            self.hist_to.set(today.strftime("%d-%m-%Y"))
        elif period == "Last Month":
            first_this = today.replace(day=1)
            last_prev = first_this - datetime.timedelta(days=1)
            first_prev = last_prev.replace(day=1)
            self.hist_from.set(first_prev.strftime("%d-%m-%Y"))
            self.hist_to.set(last_prev.strftime("%d-%m-%Y"))
        elif re.fullmatch(r"\\d{4}", period or ""):
            self.hist_from.set(f"01-01-{period}")
            self.hist_to.set(f"31-12-{period}")
        elif period == "All":
            self.hist_from.set("")
            self.hist_to.set("")
        self.refresh_history()

    def _clear_history_filters(self):
        self.hist_period.set("All")
        self.hist_from.set("")
        self.hist_to.set("")
        self.hist_search.set("")
        if hasattr(self, "hist_user"):
            self.hist_user.set("All Users")
        if hasattr(self, "hist_bill_type"):
            self.hist_bill_type.set("All Types")
        self.refresh_history()

    def _history_date_range(self):
        period = self.hist_period.get()
        if period == "All":
            return None, None
        if period == "Custom Range":
            start = self._history_parse_date(self.hist_from.get())
            end = self._history_parse_date(self.hist_to.get())
        else:
            start = self._history_parse_date(self.hist_from.get())
            end = self._history_parse_date(self.hist_to.get())
        return start, end

    def _bill_payment_values(self, bill):
        total = float(_bill_totals(bill)["grand_total"])
        try:
            paid = max(0.0, min(total, float(bill.get("amount_paid", 0) or 0)))
        except (TypeError, ValueError):
            paid = 0.0
        if total > 0 and paid >= total - 0.005:
            paid = total
            status = "Paid"
        elif paid <= 0.005:
            paid = 0.0
            status = "Not Paid"
        else:
            status = "Partial"
        return status, paid, max(0.0, total - paid)

    def refresh_history(self):
        if not hasattr(self, "hist_tree"):
            return
        for r in self.hist_tree.get_children():
            self.hist_tree.delete(r)

        q = self.hist_search.get().lower().strip()
        selected_user = self.hist_user.get().strip() if hasattr(self, "hist_user") else "All Users"
        selected_bill_type = (
            self.hist_bill_type.get().strip()
            if hasattr(self, "hist_bill_type") else "All Types"
        )
        start, end = self._history_date_range()
        shown = []
        for b in reversed(BILLS):
            bill_user = str(b.get("billing_user", "User 1") or "User 1").strip()
            bill_type = str(b.get("bill_type", "Bill Of Supply") or "Bill Of Supply").strip()
            if selected_user != "All Users" and bill_user != selected_user:
                continue
            if selected_bill_type != "All Types" and bill_type != selected_bill_type:
                continue

            bd = self._history_parse_date(b.get("date"))
            if start and (bd is None or bd < start):
                continue
            if end and (bd is None or bd > end):
                continue

            hay = (
                f"{b.get('bill_no','')} {b.get('customer','')} "
                f"{b.get('mobile','')} {b.get('address','')} {b.get('date','')} "
                f"{b.get('payment_status','')} {b.get('billing_user','')} "
                f"{b.get('cgst_pct','')} {b.get('sgst_pct','')} "
                f"{b.get('grand_total','')}"
            ).lower()
            if q and q not in hay:
                continue

            status, paid, balance = self._bill_payment_values(b)
            tag = "unpaid" if status == "Not Paid" or balance > 0 else "paid"
            self.hist_tree.insert(
                "", "end", iid=b["bill_no"], tags=(tag,),
                values=(
                    b.get("bill_no", ""), b.get("date", ""),
                    b.get("customer", "-") or "-", b.get("mobile", "-") or "-",
                    len(b.get("items", [])), f"₹ {_bill_totals(b)['grand_total']:,.2f}",
                    status, f"₹ {paid:,.2f}", f"₹ {balance:,.2f}"
                )
            )
            shown.append(b)

        total_value = sum(_bill_totals(b)["grand_total"] for b in shown)
        paid_value = sum(self._bill_payment_values(b)[1] for b in shown)
        balance_value = sum(self._bill_payment_values(b)[2] for b in shown)
        unpaid_count = sum(1 for b in shown if self._bill_payment_values(b)[2] > 0)
        if hasattr(self, "hist_summary"):
            user_text = selected_user
            type_text = selected_bill_type
            self.hist_summary.config(
                text=f"Showing {len(shown)} bills   |   User: {user_text}   |   "
                     f"Type: {type_text}   |   Sales: ₹ {total_value:,.2f}   |   "
                     f"Paid: ₹ {paid_value:,.2f}   |   Pending: ₹ {balance_value:,.2f} "
                     f"({unpaid_count} bills)"
            )

    def _history_selected_bill(self):
        sel = self.hist_tree.selection()
        if not sel:
            messagebox.showwarning("Select Bill", "Select a bill first.")
            return None
        bill_no = sel[0]
        return next((b for b in BILLS if b.get("bill_no") == bill_no), None)

    def _history_double_click(self, event=None):
        if event is not None:
            row = self.hist_tree.identify_row(event.y)
            if row:
                self.hist_tree.selection_set(row)
                self.hist_tree.focus(row)
        self.edit_history_bill()
        return "break"

    def _history_right_click(self, event):
        row = self.hist_tree.identify_row(event.y)
        if not row:
            return
        self.hist_tree.selection_set(row)
        self.hist_tree.focus(row)
        menu = tk.Menu(self, tearoff=0, bg=PANEL, fg=TEXT,
                       activebackground=ACCENT, activeforeground="#111", font=FONT)
        menu.add_command(label="🖨  Print", command=self.print_history_bill)
        menu.add_command(label="✏  Edit in New Bill", command=self.edit_history_bill)
        menu.add_command(label="📄  Open PDF", command=self.open_history_pdf)
        menu.tk_popup(event.x_root, event.y_root)

    def _history_bill_pdf(self, bill):
        stored = str(bill.get("pdf_path", "") or "").strip()
        if stored:
            return stored
        # Legacy bills created before timestamped PDF filenames are opened from
        # their old location when it still exists.
        safe = str(bill.get("bill_no", "bill")).replace("/", "_")
        legacy = os.path.join(PDF_DIR, f"{safe}.pdf")
        if os.path.exists(legacy):
            return legacy
        return self._bill_pdf_path(bill, new_file=False)

    def _open_pdf_file(self, out_pdf):
        if not os.path.exists(out_pdf):
            return False
        try:
            if os.name == "nt":
                os.startfile(out_pdf)
            else:
                os.system(f'xdg-open "{out_pdf}"')
            return True
        except Exception:
            return False

    def print_history_bill(self):
        """Generate PDF, open it in Chrome, then automatically send Ctrl+P."""
        bill = self._history_selected_bill()
        if not bill:
            return

        out_pdf = self._history_bill_pdf(bill)

        try:
            generate_a4_invoice(bill, out_pdf)

            if not os.path.exists(out_pdf):
                raise FileNotFoundError(f"PDF was not generated: {out_pdf}")

            if os.name == "nt":
                # Open the PDF in Chrome, just like opening it normally.
                # Then send Ctrl+P to the Chrome window automatically.
                chrome_candidates = [
                    os.path.join(os.environ.get("PROGRAMFILES", ""),
                                 "Google", "Chrome", "Application", "chrome.exe"),
                    os.path.join(os.environ.get("PROGRAMFILES(X86)", ""),
                                 "Google", "Chrome", "Application", "chrome.exe"),
                    os.path.join(os.environ.get("LOCALAPPDATA", ""),
                                 "Google", "Chrome", "Application", "chrome.exe"),
                ]
                chrome_exe = next(
                    (p for p in chrome_candidates if p and os.path.exists(p)),
                    None
                )

                if chrome_exe:
                    import subprocess
                    import time

                    # Use a new Chrome window so the PDF is the active document.
                    subprocess.Popen(
                        [chrome_exe, "--new-window", os.path.abspath(out_pdf)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
                    time.sleep(2.0)

                    # Windows built-in WScript.Shell sends the same Ctrl+P
                    # keyboard shortcut a user would press in Chrome.
                    pdf_title = os.path.basename(out_pdf)
                    ps_script = (
                        "$ws = New-Object -ComObject WScript.Shell; "
                        "$ok = $ws.AppActivate("
                        + repr(pdf_title)
                        + "); "
                        "if (-not $ok) { "
                        "$ok = $ws.AppActivate('Google Chrome') "
                        "}; "
                        "Start-Sleep -Milliseconds 500; "
                        "$ws.SendKeys('^p')"
                    )

                    subprocess.Popen(
                        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                         "-Command", ps_script],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
                    return

                # If Chrome cannot be found, use the normal PDF association.
                if self._open_pdf_file(out_pdf):
                    messagebox.showinfo(
                        "Print",
                        "Chrome was not found automatically. "
                        "The PDF has been opened; press Ctrl+P to print."
                    )
                    return

                raise RuntimeError("Could not open the PDF for printing.")

            # Non-Windows fallback: keep existing behavior.
            if not self._open_pdf_file(out_pdf):
                raise RuntimeError("Could not open the generated PDF.")

        except Exception as e:
            messagebox.showerror("Print Error", str(e))

    def open_history_pdf(self):
        bill = self._history_selected_bill()
        if not bill:
            return
        out_pdf = self._history_bill_pdf(bill)
        try:
            generate_a4_invoice(bill, out_pdf)
            if not self._open_pdf_file(out_pdf):
                messagebox.showwarning("PDF", "Could not open the PDF file on this system.")
        except Exception as e:
            messagebox.showerror("PDF Error", str(e))

    def reprint_bill(self):
        self.print_history_bill()

    def edit_history_bill(self):
        bill = self._history_selected_bill()
        if not bill:
            return
        self.show_billing(edit_bill=bill)

    def _close_history_editor(self):
        self._history_editor = None

    def delete_bill(self):
        bill = self._history_selected_bill()
        if not bill:
            return
        bill_no = bill.get("bill_no", "")
        if not messagebox.askyesno("Delete", f"Delete bill {bill_no}?"):
            return
        global BILLS
        # Restore stock for a deleted bill before removing the sales record.
        self._adjust_stock_for_bill_change(bill, {
            **bill,
            "items": []
        })
        BILLS = [b for b in BILLS if b.get("bill_no") != bill_no]
        save_json(BILLS_FILE, BILLS)
        self.refresh_history()

    # =========================================================
    # STOCK MANAGEMENT
    # =========================================================
    def show_stock(self):
        """Stock Management with a responsive two-pane layout.

        The stock calculations and purchase/billing links are unchanged.
        Only the geometry is reorganized so the Purchase / Stock In form,
        stock table and yearly summary remain visible on normal screens.
        """
        self.current_view = "show_stock"
        self.clear_content()
        self.stock_editing_index = None

        head = tk.Frame(self.content, bg=BG)
        head.pack(fill="x", padx=20, pady=(16, 7))
        tk.Label(head, text="📦  Stock Management", bg=BG, fg=TEXT,
                 font=FONT_H1).pack(side="left")
        tk.Label(head, text="Products ↔ Purchases ↔ Billing ↔ Profit",
                 bg=BG, fg=MUTED, font=FONT).pack(side="right")

        # KPI summary
        summary = tk.Frame(self.content, bg=BG)
        summary.pack(fill="x", padx=20, pady=(0, 7))
        self.stock_summary_labels = {}
        for key, label, color in [
            ("products", "Products", INFO),
            ("purchase_qty", "Purchased Qty", SUCCESS),
            ("purchase_value", "Purchase Value", "#8b5cf6"),
            ("sold_qty", "Billed Qty", ACCENT),
            ("sales_value", "Sales Value", "#06b6d4"),
            ("profit", "Profit", "#14b8a6"),
        ]:
            card = tk.Frame(summary, bg=PANEL)
            card.pack(side="left", expand=True, fill="x", padx=3, ipady=3)
            tk.Label(card, text=label, bg=PANEL, fg=MUTED,
                     font=("Segoe UI", 8, "bold")).pack(pady=(5, 1))
            lab = tk.Label(card, text="0", bg=PANEL, fg=color,
                           font=("Segoe UI", 14, "bold"))
            lab.pack(pady=(0, 5))
            self.stock_summary_labels[key] = lab

        # Search/filter toolbar
        top = tk.Frame(self.content, bg=PANEL)
        top.pack(fill="x", padx=20, pady=(0, 7))
        tk.Label(top, text="Search Product", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(10, 5))
        self.stock_search = tk.StringVar()
        self.stock_search.trace_add("write", lambda *_: self.refresh_stock())
        tk.Entry(
            top, textvariable=self.stock_search, bg=INPUT_BG, fg=TEXT,
            insertbackground=TEXT, relief="flat", font=FONT, width=26
        ).pack(side="left", padx=5, pady=7, ipady=5)

        tk.Label(top, text="Year", bg=PANEL, fg=MUTED,
                 font=FONT_B).pack(side="left", padx=(14, 5))
        years = self._available_stock_years()
        self.stock_year = tk.StringVar(value="All")
        self.stock_year_combo = ttk.Combobox(
            top, textvariable=self.stock_year,
            values=["All"] + years, state="readonly",
            width=10, font=FONT
        )
        self.stock_year_combo.pack(side="left", padx=5, ipady=3)
        self.stock_year_combo.bind(
            "<<ComboboxSelected>>", lambda e: self.refresh_stock()
        )
        tk.Button(
            top, text="🔄 Refresh", bg=INFO, fg="white",
            font=FONT_B, relief="flat", command=self.refresh_stock
        ).pack(side="right", padx=10, pady=5)

        # Main stock area: flexible table on the left + fixed purchase form
        # on the right. This prevents the table's wide columns from pushing
        # the purchase panel off-screen.
        main = tk.Frame(self.content, bg=BG)
        main.pack(fill="both", expand=True, padx=20, pady=(0, 7))
        main.columnconfigure(0, weight=1)
        main.columnconfigure(1, weight=0)
        main.rowconfigure(0, weight=1)

        left = tk.Frame(main, bg=PANEL)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 7))
        left.rowconfigure(1, weight=1)
        left.columnconfigure(0, weight=1)

        tk.Label(
            left, text="📊 Product Stock / Sales / Profit",
            bg=PANEL, fg=ACCENT, font=FONT_H2
        ).grid(row=0, column=0, sticky="w", padx=10, pady=(7, 4))

        wrap = tk.Frame(left, bg=PANEL)
        wrap.grid(row=1, column=0, sticky="nsew", padx=7, pady=(0, 7))
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)

        self._tree_style()
        cols = (
            "Product", "Unit", "Purchased", "Sold", "Balance", "Avg Cost",
            "Selling Price", "Sales", "Profit", "Stock Value", "Potential Profit"
        )
        self.stock_tree = ttk.Treeview(
            wrap, columns=cols, show="headings", style="Angel.Treeview"
        )
        widths = {
            "Product": 210, "Unit": 55, "Purchased": 85, "Sold": 75,
            "Balance": 80, "Avg Cost": 90, "Selling Price": 100,
            "Sales": 100, "Profit": 95, "Stock Value": 105,
            "Potential Profit": 120
        }
        for c in cols:
            self.stock_tree.heading(c, text=c)
            self.stock_tree.column(
                c, width=widths[c], minwidth=55,
                stretch=(c == "Product"),
                anchor="w" if c == "Product" else "center"
            )

        self.stock_tree.tag_configure(
            "negative", background="#fecaca", foreground="#7f1d1d"
        )
        self.stock_tree.tag_configure(
            "normal", background=INPUT_BG, foreground=TEXT
        )
        self.stock_tree.grid(row=0, column=0, sticky="nsew")

        ysb = tk.Scrollbar(wrap, orient="vertical", command=self.stock_tree.yview)
        ysb.grid(row=0, column=1, sticky="ns")
        xsb = tk.Scrollbar(wrap, orient="horizontal", command=self.stock_tree.xview)
        xsb.grid(row=1, column=0, sticky="ew")
        self.stock_tree.configure(
            yscrollcommand=ysb.set, xscrollcommand=xsb.set
        )

        self.stock_tree.bind(
            "<<TreeviewSelect>>", lambda e: self._stock_selection_changed()
        )
        self.stock_tree.bind(
            "<Double-1>", lambda e: self._stock_selection_changed()
        )
        self.stock_tree.bind("<Button-3>", self._stock_right_click)

        # Fixed-width Purchase / Stock In panel.
        right = tk.Frame(main, bg=PANEL, width=360)
        right.grid(row=0, column=1, sticky="ns")
        right.grid_propagate(False)
        self._build_stock_form(right)

        # Fixed yearly summary below the main stock area.
        year_box = tk.Frame(self.content, bg=PANEL)
        year_box.pack(fill="x", padx=20, pady=(0, 12))
        yhead = tk.Frame(year_box, bg=PANEL)
        yhead.pack(fill="x", padx=10, pady=(6, 3))
        tk.Label(
            yhead, text="📅 Yearly Purchase / Billing / Profit Summary",
            bg=PANEL, fg=ACCENT, font=FONT_H2
        ).pack(side="left")
        tk.Label(
            yhead, text="Purchase Value vs Sales Value vs Profit",
            bg=PANEL, fg=MUTED, font=("Segoe UI", 9)
        ).pack(side="right")

        year_wrap = tk.Frame(year_box, bg=PANEL)
        year_wrap.pack(fill="x", padx=8, pady=(0, 7))
        self.year_tree = ttk.Treeview(
            year_wrap,
            columns=(
                "Year", "Purchased Qty", "Purchase Value",
                "Sold Qty", "Sales Value", "Profit"
            ),
            show="headings", height=3, style="Angel.Treeview"
        )
        for c, w in [
            ("Year", 70), ("Purchased Qty", 110),
            ("Purchase Value", 130), ("Sold Qty", 95),
            ("Sales Value", 120), ("Profit", 120)
        ]:
            self.year_tree.heading(c, text=c)
            self.year_tree.column(c, width=w, anchor="center")
        self.year_tree.pack(fill="x")

        self._stock_form_clear()
        self.refresh_stock()

    def _available_stock_years(self):
        years = set()
        for s in STOCK:
            for p in s.get("purchases", []) or []:
                y = self._year_from_date(p.get("date", ""))
                if y:
                    years.add(y)
            y = self._year_from_date(s.get("last_purchase_date", ""))
            if y:
                years.add(y)
        for b in BILLS:
            y = self._year_from_date(b.get("date", ""))
            if y:
                years.add(y)
        return sorted(years, reverse=True)

    def _year_from_date(self, value):
        text = str(value or "").strip()
        for fmt in ("%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y"):
            try:
                return datetime.datetime.strptime(text, fmt).strftime("%Y")
            except ValueError:
                pass
        m = re.search(r"(20\d{2})", text)
        return m.group(1) if m else ""

    def _stock_product_for_name(self, name):
        n = str(name or "").strip().lower()
        return next((p for p in PRODUCTS if str(p.get("name", "")).strip().lower() == n), None)

    def _stock_find(self, product_id=None, product_name=None):
        pid = str(product_id or "").strip()
        name = str(product_name or "").strip().lower()
        for s in STOCK:
            if pid and str(s.get("product_id", "")) == pid:
                return s
            if name and str(s.get("product_name", "")).strip().lower() == name:
                return s
        return None

    def _stock_purchase_data(self, stock):
        purchases = stock.get("purchases", []) or []
        if purchases:
            qty = sum(float(p.get("qty", 0) or 0) for p in purchases)
            value = sum(float(p.get("qty", 0) or 0) * float(p.get("rate", 0) or 0) for p in purchases)
        else:
            qty = float(stock.get("purchased_qty", 0) or 0)
            value = qty * float(stock.get("purchase_rate", 0) or 0)
        avg = value / qty if qty > 0 else float(stock.get("purchase_rate", 0) or 0)
        return qty, value, avg

    def _stock_bill_metrics(self, product):
        pid = str(product.get("id", "")).strip()
        name = str(product.get("name", "")).strip().lower()
        sold = sales = 0.0
        for b in BILLS:
            for it in b.get("items", []) or []:
                iid = str(it.get("id", "")).strip()
                iname = str(it.get("name", "")).strip().lower()
                matched = bool(pid and iid and iid == pid) or (not iid and iname == name)
                if not matched and not iid and iname == name:
                    matched = True
                if matched:
                    q = float(it.get("qty", 0) or 0)
                    r = float(it.get("rate", 0) or 0)
                    sold += q
                    sales += q * r
        return sold, sales

    def _stock_aggregate_rows(self):
        rows = []
        seen = set()
        for p in PRODUCTS:
            pid = str(p.get("id", ""))
            name = str(p.get("name", ""))
            stock = self._stock_find(pid, name) or {}
            purchased, purchase_value, avg_cost = self._stock_purchase_data(stock)
            sold, sales = self._stock_bill_metrics(p)
            balance = purchased - sold
            sell_price = float(p.get("discount", p.get("rate", 0)) or 0)
            cost_sold = sold * avg_cost
            profit = sales - cost_sold
            stock_value = max(0.0, balance) * avg_cost
            potential = max(0.0, balance) * (sell_price - avg_cost)
            rows.append({
                "product": p, "stock": stock, "purchased": purchased,
                "purchase_value": purchase_value, "sold": sold, "sales": sales,
                "balance": balance, "avg_cost": avg_cost, "sell_price": sell_price,
                "cost_sold": cost_sold, "profit": profit,
                "stock_value": stock_value, "potential_profit": potential,
            })
            seen.add(pid or name.lower())

        # Preserve legacy/orphan stock records too, but link them where possible.
        for s in STOCK:
            key = str(s.get("product_id") or s.get("product_name") or "").strip().lower()
            if not key or key in {x for x in seen}:
                continue
            purchased, purchase_value, avg_cost = self._stock_purchase_data(s)
            name = str(s.get("product_name", ""))
            sold = float(s.get("sold_qty", 0) or 0)
            sales = sold * float(s.get("selling_price", 0) or 0)
            balance = purchased - sold
            rows.append({
                "product": {"id": s.get("product_id", ""), "name": name, "unit": s.get("unit", "Box"), "discount": s.get("selling_price", 0)},
                "stock": s, "purchased": purchased, "purchase_value": purchase_value,
                "sold": sold, "sales": sales, "balance": balance, "avg_cost": avg_cost,
                "sell_price": float(s.get("selling_price", 0) or 0),
                "cost_sold": sold * avg_cost, "profit": sales - sold * avg_cost,
                "stock_value": max(0.0, balance) * avg_cost,
                "potential_profit": max(0.0, balance) * (float(s.get("selling_price", 0) or 0) - avg_cost),
            })
        return rows

    def _build_stock_form(self, parent):
        tk.Label(parent, text="➕ Purchase / Stock In", bg=PANEL, fg=ACCENT,
                 font=FONT_H2).pack(anchor="w", padx=12, pady=(10, 5))
        self.stock_form_status = tk.Label(parent, text="", bg=PANEL, fg=MUTED,
                                          font=("Segoe UI", 9), wraplength=335, justify="left")
        self.stock_form_status.pack(anchor="w", padx=12, pady=(0, 5))

        def label(text):
            tk.Label(parent, text=text, bg=PANEL, fg=MUTED,
                     font=("Segoe UI", 9, "bold")).pack(anchor="w", padx=12, pady=(5, 1))

        label("Product (linked to Products master)")
        self.stock_product_var = tk.StringVar()
        self.stock_product_combo = ttk.Combobox(
            parent, textvariable=self.stock_product_var,
            values=[p.get("name", "") for p in PRODUCTS],
            state="normal", font=FONT
        )
        self.stock_product_combo.pack(fill="x", padx=12, ipady=5)
        self.stock_product_combo.bind("<<ComboboxSelected>>", lambda e: self._stock_product_changed())
        self.stock_product_combo.bind("<KeyRelease>", lambda e: self._stock_product_changed())

        label("Unit")
        self.stock_unit_var = tk.StringVar(value="Box")
        tk.Entry(parent, textvariable=self.stock_unit_var, bg=ROWBG, fg=SUCCESS,
                 relief="flat", font=FONT_B, state="readonly").pack(fill="x", padx=12, ipady=6)

        label("Purchase Quantity")
        self.stock_qty_var = tk.StringVar()
        tk.Entry(parent, textvariable=self.stock_qty_var, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT).pack(fill="x", padx=12, ipady=6)

        label("Purchase Rate ₹")
        self.stock_rate_var = tk.StringVar()
        tk.Entry(parent, textvariable=self.stock_rate_var, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT).pack(fill="x", padx=12, ipady=6)

        label("Supplier")
        self.stock_supplier_var = tk.StringVar()
        tk.Entry(parent, textvariable=self.stock_supplier_var, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT).pack(fill="x", padx=12, ipady=6)

        label("Purchase Date (DD-MM-YYYY)")
        self.stock_date_var = tk.StringVar(value=datetime.datetime.now().strftime("%d-%m-%Y"))
        tk.Entry(parent, textvariable=self.stock_date_var, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT).pack(fill="x", padx=12, ipady=6)

        label("Supplier Invoice / Purchase Bill No.")
        self.stock_invoice_var = tk.StringVar()
        tk.Entry(parent, textvariable=self.stock_invoice_var, bg=INPUT_BG, fg=TEXT,
                 insertbackground=TEXT, relief="flat", font=FONT).pack(fill="x", padx=12, ipady=6)

        self.stock_current_label = tk.Label(parent, text="Current Stock: 0   |   Selling Price: ₹0.00",
                                             bg=PANEL, fg=ACCENT, font=FONT_B)
        self.stock_current_label.pack(anchor="w", padx=12, pady=(8, 4))

        btn = tk.Frame(parent, bg=PANEL)
        btn.pack(fill="x", padx=12, pady=(3, 6))
        tk.Button(btn, text="💾 Save Purchase", bg=SUCCESS, fg="white",
                  font=FONT_B, relief="flat", command=self._stock_save_form
                  ).pack(side="left", expand=True, fill="x", padx=(0, 4), ipady=7)
        tk.Button(btn, text="↺ Clear", bg=INFO, fg="white",
                  font=FONT_B, relief="flat", command=self._stock_form_clear
                  ).pack(side="right", expand=True, fill="x", padx=(4, 0), ipady=7)

        tk.Label(parent, text="🧾 Purchase History", bg=PANEL, fg=ACCENT,
                 font=FONT_B).pack(anchor="w", padx=12, pady=(8, 3))
        hw = tk.Frame(parent, bg=PANEL)
        hw.pack(fill="both", expand=True, padx=12, pady=(0, 10))
        hcols = ("Date", "Qty", "Rate", "Supplier", "Invoice")
        self.stock_purchase_tree = ttk.Treeview(hw, columns=hcols, show="headings",
                                                style="Angel.Treeview", height=7)
        for c, w in [("Date",82),("Qty",55),("Rate",70),("Supplier",100),("Invoice",95)]:
            self.stock_purchase_tree.heading(c, text=c)
            self.stock_purchase_tree.column(c, width=w, anchor="center")
        self.stock_purchase_tree.pack(side="left", fill="both", expand=True)
        hs = tk.Scrollbar(hw, command=self.stock_purchase_tree.yview)
        self.stock_purchase_tree.configure(yscrollcommand=hs.set)
        hs.pack(side="right", fill="y")

    def _stock_product_changed(self):
        typed = self.stock_product_var.get().strip()
        names = [str(p.get("name", "")) for p in PRODUCTS]
        if hasattr(self, "stock_product_combo"):
            q = typed.lower()
            filtered = [n for n in names if q in n.lower()] if q else names
            self.stock_product_combo["values"] = filtered[:200]
        prod = self._stock_product_for_name(typed)
        if prod:
            self.stock_unit_var.set(prod.get("unit", "Box"))
            self.stock_rate_var.set(str(float(prod.get("rate", 0) or 0)))
            stock = self._stock_find(prod.get("id", ""), prod.get("name", ""))
            purchased, _, _ = self._stock_purchase_data(stock or {})
            sold, _ = self._stock_bill_metrics(prod)
            balance = purchased - sold
            sell = float(prod.get("discount", prod.get("rate", 0)) or 0)
            self.stock_current_label.config(text=f"Current Stock: {balance:g}   |   Selling Price: ₹{sell:,.2f}")
            self._refresh_stock_purchase_history(stock)
        else:
            self.stock_current_label.config(text="Current Stock: 0   |   Selling Price: ₹0.00")
            self._refresh_stock_purchase_history(None)

    def _stock_selection_changed(self):
        sel = self.stock_tree.selection() if hasattr(self, "stock_tree") else ()
        if not sel:
            return
        try:
            row = self._stock_aggregate_rows()[int(sel[0])]
        except (ValueError, IndexError):
            return
        p = row["product"]
        self.stock_editing_index = None
        self.stock_product_var.set(str(p.get("name", "")))
        self._stock_product_changed()
        stock = row.get("stock") or {}
        self.stock_supplier_var.set(stock.get("supplier", ""))
        self.stock_date_var.set(stock.get("last_purchase_date", datetime.datetime.now().strftime("%d-%m-%Y")))
        self.stock_invoice_var.set(stock.get("supplier_invoice", ""))
        self.stock_qty_var.set("")
        self.stock_rate_var.set(str(float(row.get("avg_cost", 0) or 0)))
        self.stock_form_status.config(text=f"Selected: {p.get('name','')} — use Save Purchase to add new stock.")

    def _stock_form_clear(self):
        if not hasattr(self, "stock_product_var"):
            return
        self.stock_editing_index = None
        self.stock_product_var.set("")
        self.stock_unit_var.set("Box")
        self.stock_qty_var.set("")
        self.stock_rate_var.set("")
        self.stock_supplier_var.set("")
        self.stock_date_var.set(datetime.datetime.now().strftime("%d-%m-%Y"))
        self.stock_invoice_var.set("")
        self.stock_current_label.config(text="Current Stock: 0   |   Selling Price: ₹0.00")
        self.stock_form_status.config(text="Type/select a product. Product data is linked directly to Products master.")
        self._refresh_stock_purchase_history(None)

    def _stock_save_form(self):
        name = self.stock_product_var.get().strip()
        prod = self._stock_product_for_name(name)
        if not prod:
            self.stock_form_status.config(text="❌ Product must be selected from the Products master.", fg=DANGER)
            return
        try:
            qty = float(self.stock_qty_var.get())
            rate = float(self.stock_rate_var.get())
            if qty <= 0 or rate < 0:
                raise ValueError
        except ValueError:
            self.stock_form_status.config(text="❌ Purchase quantity must be > 0 and rate must be valid.", fg=DANGER)
            return

        stock = self._stock_find(prod.get("id", ""), name)
        if stock is None:
            stock = {
                "product_id": prod.get("id", ""),
                "product_name": prod.get("name", ""),
                "unit": prod.get("unit", "Box"),
                "purchased_qty": 0.0, "sold_qty": 0.0,
                "purchase_rate": rate, "selling_price": float(prod.get("discount", prod.get("rate", 0)) or 0),
                "supplier": "", "last_purchase_date": "", "supplier_invoice": "",
                "last_purchase_qty": 0.0, "purchases": [],
            }
            STOCK.append(stock)

        stock["product_id"] = prod.get("id", "")
        stock["product_name"] = prod.get("name", "")
        stock["unit"] = prod.get("unit", "Box")
        stock["purchase_rate"] = rate
        stock["selling_price"] = float(prod.get("discount", prod.get("rate", 0)) or 0)
        stock["supplier"] = self.stock_supplier_var.get().strip()
        stock["last_purchase_date"] = self.stock_date_var.get().strip()
        stock["supplier_invoice"] = self.stock_invoice_var.get().strip()
        stock["last_purchase_qty"] = qty
        stock.setdefault("purchases", []).append({
            "date": self.stock_date_var.get().strip(),
            "qty": qty, "rate": rate,
            "supplier": self.stock_supplier_var.get().strip(),
            "invoice": self.stock_invoice_var.get().strip(),
        })
        # Keep legacy aggregate fields synchronized.
        stock["purchased_qty"] = sum(float(p.get("qty", 0) or 0) for p in stock.get("purchases", []))
        stock["sold_qty"] = self._stock_bill_metrics(prod)[0]

        save_json(STOCK_FILE, STOCK)
        self.stock_year_combo["values"] = ["All"] + self._available_stock_years()
        self.stock_form_status.config(text=f"✓ Purchase saved for {prod.get('name','')}. Billing will reduce its balance automatically.", fg=SUCCESS)
        self.refresh_stock()
        self.stock_product_var.set(prod.get("name", ""))
        self._stock_product_changed()
        self.stock_qty_var.set("")

    def _refresh_stock_purchase_history(self, stock):
        if not hasattr(self, "stock_purchase_tree"):
            return
        for r in self.stock_purchase_tree.get_children():
            self.stock_purchase_tree.delete(r)
        if not stock:
            return
        for i, p in enumerate(reversed(stock.get("purchases", []) or [])):
            self.stock_purchase_tree.insert("", "end", iid=str(i), values=(
                p.get("date", ""), f"{float(p.get('qty', 0) or 0):g}",
                f"₹ {float(p.get('rate', 0) or 0):,.2f}",
                p.get("supplier", "") or "-", p.get("invoice", "") or "-"
            ))

    def _stock_edit_form(self):
        # Same main GUI: editing selected product loads the purchase fields for the next stock entry.
        self._stock_selection_changed()

    def _stock_delete(self):
        sel = self.stock_tree.selection()
        if not sel:
            messagebox.showwarning("Select", "Select a stock/product row first.")
            return
        rows = self._stock_aggregate_rows()
        try:
            row = rows[int(sel[0])]
        except (ValueError, IndexError):
            return
        stock = row.get("stock")
        if not stock:
            messagebox.showwarning("No Purchase Record", "This product has no stock purchase record to delete.")
            return
        if not messagebox.askyesno("Delete Stock", f"Delete all purchase history for {row['product'].get('name','')}?"):
            return
        if stock in STOCK:
            STOCK.remove(stock)
            save_json(STOCK_FILE, STOCK)
        self.refresh_stock()
        self._stock_form_clear()

    def _stock_right_click(self, event):
        row = self.stock_tree.identify_row(event.y)
        if not row:
            return
        self.stock_tree.selection_set(row)
        self.stock_tree.focus(row)
        menu = tk.Menu(
            self, tearoff=0, bg=PANEL, fg=TEXT,
            activebackground=ACCENT, activeforeground="#111", font=FONT
        )
        menu.add_command(label="➕  Purchase / Stock In", command=self._stock_selection_changed)
        menu.add_command(label="✏  Edit Selected", command=self._stock_edit_form)
        menu.add_command(label="🗑  Delete Purchase History", command=self._stock_delete)
        menu.tk_popup(event.x_root, event.y_root)

    def _update_stock_from_bill(self, bill):
        # Stock is now calculated live from BILLS, so no manual decrement is needed.
        # Keep cached sold_qty for backward compatibility with older stock.json files.
        changed = False
        for it in bill.get("items", []):
            pid = str(it.get("id", "")).strip()
            name = str(it.get("name", "")).strip()
            stock = self._stock_find(pid, name)
            if stock is not None:
                stock["sold_qty"] = self._stock_bill_metrics({
                    "id": stock.get("product_id", ""), "name": stock.get("product_name", "")
                })[0]
                changed = True
        if changed:
            save_json(STOCK_FILE, STOCK)

    def _adjust_stock_for_bill_change(self, old_bill, new_bill):
        # Live stock uses BILLS as the source of truth. This method only refreshes
        # legacy sold_qty cache so existing stock.json remains compatible.
        touched = set()
        for it in (old_bill.get("items", []) or []) + (new_bill.get("items", []) or []):
            touched.add((str(it.get("id", "")).strip(), str(it.get("name", "")).strip()))
        changed = False
        for pid, name in touched:
            stock = self._stock_find(pid, name)
            if stock is not None:
                prod = self._stock_product_for_name(stock.get("product_name", "")) or {
                    "id": stock.get("product_id", ""), "name": stock.get("product_name", "")
                }
                stock["sold_qty"] = self._stock_bill_metrics(prod)[0]
                changed = True
        if changed:
            save_json(STOCK_FILE, STOCK)

    def refresh_stock(self):
        if not hasattr(self, "stock_tree"):
            return
        for r in self.stock_tree.get_children():
            self.stock_tree.delete(r)

        q = self.stock_search.get().strip().lower() if hasattr(self, "stock_search") else ""
        selected_year = self.stock_year.get() if hasattr(self, "stock_year") else "All"
        rows = self._stock_aggregate_rows()
        visible = []
        for row in rows:
            p = row["product"]
            hay = f"{p.get('id','')} {p.get('name','')} {p.get('unit','')}".lower()
            if q and q not in hay:
                continue
            # For a selected year, show products with either purchases or bills in that year.
            if selected_year != "All":
                has_year = False
                s = row.get("stock") or {}
                for pur in s.get("purchases", []) or []:
                    if self._year_from_date(pur.get("date", "")) == selected_year:
                        has_year = True; break
                if not has_year:
                    for b in BILLS:
                        if self._year_from_date(b.get("date", "")) != selected_year:
                            continue
                        for it in b.get("items", []) or []:
                            if str(it.get("id", "")).strip() == str(p.get("id", "")).strip() or (not it.get("id") and str(it.get("name", "")).strip().lower() == str(p.get("name", "")).strip().lower()):
                                has_year = True; break
                        if has_year: break
                if not has_year:
                    continue
            visible.append(row)

        for idx, row in enumerate(visible):
            p = row["product"]
            tag = "negative" if row["balance"] < 0 else "normal"
            self.stock_tree.insert("", "end", iid=str(idx), tags=(tag,), values=(
                p.get("name", ""), p.get("unit", "Box"),
                f"{row['purchased']:g}", f"{row['sold']:g}", f"{row['balance']:g}",
                f"₹ {row['avg_cost']:,.2f}", f"₹ {row['sell_price']:,.2f}",
                f"₹ {row['sales']:,.2f}", f"₹ {row['profit']:,.2f}",
                f"₹ {row['stock_value']:,.2f}", f"₹ {row['potential_profit']:,.2f}"
            ))

        # Overall live summary.
        all_rows = self._stock_aggregate_rows()
        total_pq = sum(r["purchased"] for r in all_rows)
        total_pv = sum(r["purchase_value"] for r in all_rows)
        total_sq = sum(r["sold"] for r in all_rows)
        total_sv = sum(r["sales"] for r in all_rows)
        total_profit = sum(r["profit"] for r in all_rows)
        self.stock_summary_labels["products"].config(text=str(len(PRODUCTS)))
        self.stock_summary_labels["purchase_qty"].config(text=f"{total_pq:g}")
        self.stock_summary_labels["purchase_value"].config(text=f"₹ {total_pv:,.0f}")
        self.stock_summary_labels["sold_qty"].config(text=f"{total_sq:g}")
        self.stock_summary_labels["sales_value"].config(text=f"₹ {total_sv:,.0f}")
        self.stock_summary_labels["profit"].config(text=f"₹ {total_profit:,.0f}")

        # Yearly summary.
        if hasattr(self, "year_tree"):
            for r in self.year_tree.get_children():
                self.year_tree.delete(r)
            summary = {}
            for s in STOCK:
                for pur in s.get("purchases", []) or []:
                    y = self._year_from_date(pur.get("date", "")) or "Unknown"
                    rec = summary.setdefault(y, {"pq":0.0,"pv":0.0,"sq":0.0,"sv":0.0})
                    qty = float(pur.get("qty", 0) or 0); rate = float(pur.get("rate", 0) or 0)
                    rec["pq"] += qty; rec["pv"] += qty * rate
            for b in BILLS:
                y = self._year_from_date(b.get("date", "")) or "Unknown"
                rec = summary.setdefault(y, {"pq":0.0,"pv":0.0,"sq":0.0,"sv":0.0})
                for it in b.get("items", []) or []:
                    qv = float(it.get("qty", 0) or 0); rv = float(it.get("rate", 0) or 0)
                    rec["sq"] += qv; rec["sv"] += qv * rv
            for y in sorted(summary, reverse=True):
                rec = summary[y]
                # Weighted average cost across all purchases for the year is used for an informative yearly profit estimate.
                avg = rec["pv"] / rec["pq"] if rec["pq"] > 0 else 0
                profit = rec["sv"] - rec["sq"] * avg
                self.year_tree.insert("", "end", values=(
                    y, f"{rec['pq']:g}", f"₹ {rec['pv']:,.2f}",
                    f"{rec['sq']:g}", f"₹ {rec['sv']:,.2f}", f"₹ {profit:,.2f}"
                ))

        # Keep purchase history synced to current product selection.
        self._stock_product_changed()

    # =========================================================
    # SETTINGS
    # =========================================================
    def show_settings(self):
        self.current_view = "show_settings"
        self.clear_content()
        head = tk.Frame(self.content, bg=BG)
        head.pack(fill="x", padx=20, pady=(20, 10))
        tk.Label(head, text="⚙️  Settings", bg=BG, fg=TEXT,
                 font=FONT_H1).pack(side="left")

        card = tk.Frame(self.content, bg=PANEL)
        card.pack(fill="x", padx=20, pady=10)

        def field(label, key, is_num=False):
            row = tk.Frame(card, bg=PANEL); row.pack(fill="x", padx=16, pady=6)
            tk.Label(row, text=label, bg=PANEL, fg=MUTED,
                     font=("Segoe UI", 9, "bold"), width=20, anchor="w").pack(side="left")
            v = tk.StringVar(value=str(SETTINGS.get(key, "")))
            tk.Entry(row, textvariable=v, bg=INPUT_BG, fg=TEXT,
                     insertbackground=TEXT, relief="flat",
                     font=("Segoe UI", 11)).pack(side="left", fill="x",
                                                 expand=True, ipady=6)
            return v, key, is_num

        self.setting_fields = [
            field("Shop Name", "shop_name"),
            field("Address", "address"),
            field("Phone", "phone"),
            field("Email", "email"),
            field("GSTIN", "gstin"),
            field("State", "state"),
            field("Tagline", "tagline"),
            field("Bill Prefix", "bill_prefix"),
            field("Next Bill No", "next_bill", True),
            field("Rate Edit Password", "rate_edit_password"),
        ]

        theme_row = tk.Frame(card, bg=PANEL)
        theme_row.pack(fill="x", padx=16, pady=(10, 12))
        tk.Label(theme_row, text="Application Theme", bg=PANEL, fg=MUTED,
                 font=("Segoe UI", 9, "bold"), width=20, anchor="w").pack(side="left")
        self.theme_var = tk.StringVar(value=SETTINGS.get("theme", "Dark Blue"))
        self.theme_combo = ttk.Combobox(
            theme_row, textvariable=self.theme_var,
            values=list(THEMES.keys()), state="readonly",
            font=("Segoe UI", 11), width=22
        )
        self.theme_combo.pack(side="left", padx=4, ipady=5)
        tk.Label(theme_row, text="Dark Blue / Light / Midnight",
                 bg=PANEL, fg=MUTED, font=("Segoe UI", 9)).pack(side="left", padx=10)

        tk.Button(self.content, text="💾 Save Settings",
                  bg=SUCCESS, fg="white", font=FONT_B, relief="flat",
                  padx=20, pady=10, command=self.save_settings).pack(pady=20)

        tk.Label(self.content,
                 text=f"Data folder: {DATA_DIR}",
                 bg=BG, fg=MUTED, font=("Segoe UI", 9)).pack(pady=(4, 0))

    def save_settings(self):
        for var, key, is_num in self.setting_fields:
            val = var.get().strip()
            if is_num:
                try: val = int(val)
                except ValueError: val = 1
            SETTINGS[key] = val

        new_theme = self.theme_var.get().strip() or "Dark Blue"
        SETTINGS["theme"] = new_theme
        save_json(SETTINGS_FILE, SETTINGS)

        _apply_theme(new_theme)
        self.configure(bg=BG)
        self.sidebar.configure(bg=PANEL)
        self.content.configure(bg=BG)
        messagebox.showinfo("Saved", f"Settings saved. Theme: {new_theme}")
        # Rebuild the same page so every widget immediately receives the new theme.
        current = getattr(self, "current_view", "show_settings")
        getattr(self, current)()

    # =========================================================
    # HELPERS
    # =========================================================
    def _select_all_text(self, event=None):
        widget = event.widget if event is not None else self.focus_get()
        if isinstance(widget, tk.Text):
            try:
                widget.tag_add("sel", "1.0", "end-1c")
                widget.mark_set("insert", "end-1c")
                widget.see("insert")
                return "break"
            except Exception:
                pass
        if isinstance(widget, (tk.Entry, ttk.Entry)):
            try:
                widget.select_range(0, tk.END)
                widget.icursor(tk.END)
                return "break"
            except Exception:
                pass
        return None

    def _tree_style(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Angel.Treeview",
                        background=INPUT_BG, fieldbackground=INPUT_BG,
                        foreground=TEXT, rowheight=28,
                        font=("Segoe UI", 10), borderwidth=0)
        style.configure("Angel.Treeview.Heading",
                        background=PANEL, foreground=ACCENT,
                        font=("Segoe UI", 10, "bold"), relief="flat")
        style.map("Angel.Treeview",
                  background=[("selected", ACCENT)],
                  foreground=[("selected", "#111")])

# =========================================================
# RUN
# =========================================================
if __name__ == "__main__":
    AngelApp().mainloop()