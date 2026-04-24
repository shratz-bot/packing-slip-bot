#!/usr/bin/env python3
"""
eBay Packing Slip Generator
Reads unprocessed orders from Google Sheet, generates PDF packing slips,
uploads them to Google Drive, and writes the file URL back to the sheet.

Requirements:
    pip install gspread google-auth google-api-python-client weasyprint

Environment variables (set as GitHub Secrets):
    GOOGLE_SERVICE_ACCOUNT_JSON  - Full JSON content of the service account key
    SPREADSHEET_ID               - Google Sheet ID
    DRIVE_FOLDER_ID              - Google Drive folder ID for packing slips
    EBAY_LOGO_FILE_ID            - Google Drive file ID for the eBay logo PNG
    SHRATZ_LOGO_FILE_ID          - Google Drive file ID for the Shratz105 logo PNG
    WHATNOT_QR_FILE_ID           - Google Drive file ID for the Whatnot QR code PNG
"""

import os
import io
import json
import base64
import logging
import tempfile
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload, MediaIoBaseDownload
from weasyprint import HTML

# ── Logging ────────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)
log = logging.getLogger(__name__)

# ── Config ─────────────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SHEET_NAME       = "Sheet1"
SPREADSHEET_ID   = os.environ["SPREADSHEET_ID"]
DRIVE_FOLDER_ID  = os.environ["DRIVE_FOLDER_ID"]
EBAY_LOGO_ID     = os.environ["EBAY_LOGO_FILE_ID"]
SHRATZ_LOGO_ID   = os.environ["SHRATZ_LOGO_FILE_ID"]
WHATNOT_QR_ID    = os.environ["WHATNOT_QR_FILE_ID"]

# Column indexes (1-based, matching spreadsheet headers)
COL = {
    "ORDER_ID":         1,
    "ORDER_DATE":       2,
    "SALES_RECORD":     3,
    "BUYER_USERNAME":   4,
    "BUYER_NAME":       5,
    "ADDR_LINE1":       6,
    "ADDR_CITY":        7,
    "ADDR_STATE":       8,
    "ADDR_ZIP":         9,
    "ADDR_COUNTRY":     10,
    "BUYER_PHONE":      11,
    "BUYER_EMAIL":      12,
    "ITEM_TITLE":       13,
    "ITEM_ID":          14,
    "QUANTITY":         15,
    "ITEM_PRICE":       16,
    "SHIPPING_SERVICE": 17,
    "SHIPPING_COST":    18,
    "SHIP_BY_DATE":     19,
    "SUBTOTAL":         20,
    "SALES_TAX":        21,
    "ORDER_TOTAL":      22,
    "SLIP_GENERATED":   23,
    "DRIVE_URL":        24,
}

SELLER = {
    "name":    "Eric Schwartz",
    "line1":   "1575 Liberty Ave SE",
    "city":    "Atlanta",
    "state":   "GA",
    "zip":     "30317-2308",
    "country": "United States",
    "message": (
        "Thank you for your purchase!<br>"
        "Please reach out to me if you have any questions or concerns.<br><br>"
        "You can find me on Instagram, eBay, and Whatnot at @shratz105."
    ),
}

SHIPPING_SERVICE_MAP = {
    "USPSParcel":             "USPS Ground Advantage",
    "USPSFirstClass":         "USPS First Class",
    "USPSPriorityMail":       "USPS Priority Mail",
    "USPSPriorityMailExp":    "USPS Priority Mail Express",
    "US_eBayStandardEnvelope":"eBay Standard Envelope",
    "UPSGround":              "UPS Ground",
    "FedExGround":            "FedEx Ground",
}


# ── Google Auth ────────────────────────────────────────────────────────────────
def get_credentials():
    sa_json = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    info = json.loads(sa_json)
    return Credentials.from_service_account_info(info, scopes=SCOPES)


# ── Drive helpers ──────────────────────────────────────────────────────────────
def fetch_image_as_b64(drive_service, file_id: str) -> str:
    """Download a Drive file and return it as a base64 string."""
    request = drive_service.files().get_media(fileId=file_id)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return base64.b64encode(buf.getvalue()).decode()


def upload_pdf(drive_service, pdf_bytes: bytes, filename: str, folder_id: str) -> str:
    """Upload a PDF to Drive and return its web view URL."""
    media = MediaIoBaseUpload(
        io.BytesIO(pdf_bytes),
        mimetype="application/pdf",
        resumable=False,
    )
    file_meta = {
        "name": filename,
        "parents": [folder_id],
        "mimeType": "application/pdf",
    }
    result = drive_service.files().create(
        body=file_meta,
        media_body=media,
        fields="id, webViewLink",
    ).execute()
    return result["webViewLink"]


# ── Data helpers ───────────────────────────────────────────────────────────────
def cell(row, col_name):
    """Return a cell value by column name (1-based index)."""
    idx = COL[col_name] - 1
    return str(row[idx]).strip() if idx < len(row) else ""


def fmt_currency(val):
    try:
        return f"${float(val):.2f}"
    except (ValueError, TypeError):
        return "$0.00"


def fmt_date(val):
    """Convert various date formats to 'Mon DD, YYYY'."""
    if not val:
        return ""
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(str(val).strip(), fmt).strftime("%b %d, %Y")
        except ValueError:
            continue
    return str(val)


def fmt_phone(val):
    digits = "".join(c for c in str(val) if c.isdigit())
    if len(digits) == 10:
        return f"+1 {digits[:3]}-{digits[3:6]}-{digits[6:]}"
    if len(digits) == 11 and digits[0] == "1":
        return f"+{digits[0]} {digits[1:4]}-{digits[4:7]}-{digits[7:]}"
    return val


def fmt_shipping(val):
    return SHIPPING_SERVICE_MAP.get(val, val)


# ── HTML template ──────────────────────────────────────────────────────────────
def build_html(order: dict, ebay_b64: str, shratz_b64: str, qr_b64: str) -> str:
    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{
    font-family: Arial, Helvetica, sans-serif;
    font-size: 13px;
    color: #111;
    padding: 36px 40px;
    width: 750px;
  }}
  .slip-top {{
    display: flex;
    justify-content: space-between;
    align-items: center;
    padding-bottom: 14px;
    border-bottom: 1.5px solid #111;
    margin-bottom: 20px;
  }}
  .top-left  {{ flex: 1; display: flex; align-items: center; }}
  .top-center {{ flex: 1; display: flex; justify-content: center; }}
  .top-right {{ flex: 1; display: flex; justify-content: flex-end; align-items: flex-start; }}
  .ebay-logo   {{ height: 38px; width: auto; }}
  .shratz-logo {{ height: 66px; width: auto; }}
  .ps-label {{ font-size: 11px; font-weight: bold; letter-spacing: 0.08em; }}
  .addr-row {{ display: flex; gap: 40px; margin-bottom: 16px; }}
  .addr-block {{ flex: 1; }}
  .lbl {{ font-size: 11px; font-weight: bold; margin-bottom: 4px; }}
  .addr-block p {{ line-height: 1.6; }}
  .buyer-info {{ font-size: 12px; line-height: 1.8; margin-bottom: 16px; }}
  .order-row {{ display: flex; justify-content: space-between; align-items: baseline; margin-bottom: 4px; }}
  .order-num {{ font-size: 16px; font-weight: bold; }}
  .order-date {{ font-size: 13px; color: #555; }}
  .sales-rec {{ font-size: 12px; color: #555; margin-bottom: 12px; }}
  hr {{ border: none; border-top: 1px solid #ccc; margin: 10px 0; }}
  table {{ width: 100%; border-collapse: collapse; margin-bottom: 12px; }}
  th {{
    font-size: 11px; font-weight: bold;
    padding: 6px 0; border-bottom: 1px solid #111;
    text-align: left;
  }}
  th.r {{ text-align: right; }}
  td {{ padding: 10px 0; border-bottom: 1px solid #ddd; vertical-align: top; line-height: 1.5; }}
  td.r {{ text-align: right; }}
  .item-id {{ font-size: 12px; color: #555; }}
  .ship-note {{ font-size: 12px; font-weight: bold; text-align: right; margin-bottom: 8px; }}
  .bottom {{ display: flex; gap: 24px; margin-top: 8px; align-items: center; }}
  .msg {{ flex: 1; font-size: 12px; line-height: 1.7; }}
  .msg-title {{ font-weight: bold; margin-bottom: 4px; font-size: 12px; }}
  .qr {{ display: flex; flex-direction: column; align-items: center; gap: 4px; flex-shrink: 0; }}
  .qr-txt {{ font-size: 11px; font-weight: bold; text-align: center; }}
  .qr img {{ width: 88px; height: 88px; }}
  .totals {{ flex: 1; font-size: 13px; }}
  .trow {{ display: flex; justify-content: space-between; padding: 2px 0; }}
  .trow.grand {{
    font-weight: bold;
    border-top: 1px solid #ccc;
    margin-top: 4px;
    padding-top: 5px;
  }}
  .footnote {{
    font-size: 11px; color: #555;
    margin-top: 16px;
    border-top: 0.5px solid #ddd;
    padding-top: 10px;
    line-height: 1.6;
  }}
</style>
</head>
<body>

<div class="slip-top">
  <div class="top-left">
    <img class="ebay-logo" src="data:image/png;base64,{ebay_b64}" alt="eBay">
  </div>
  <div class="top-center">
    <img class="shratz-logo" src="data:image/jpeg;base64,{shratz_b64}" alt="Shratz105">
  </div>
  <div class="top-right">
    <div class="ps-label">PACKING SLIP</div>
  </div>
</div>

<div class="addr-row">
  <div class="addr-block">
    <div class="lbl">Ship to</div>
    <p>
      <strong>{order["buyer_name"]}</strong><br>
      {order["addr_line1"]}<br>
      {order["addr_city"]}, {order["addr_state"]}, {order["addr_zip"]}<br>
      {order["addr_country"]}
    </p>
  </div>
  <div class="addr-block">
    <div class="lbl">Ship from</div>
    <p>
      <strong>{SELLER["name"]}</strong><br>
      {SELLER["line1"]}<br>
      {SELLER["city"]}, {SELLER["state"]}, {SELLER["zip"]}<br>
      {SELLER["country"]}
    </p>
  </div>
</div>

<div class="buyer-info">
  {order["buyer_phone"]}<br>
  {order["buyer_email"]}<br>
  Username: {order["buyer_username"]}
</div>

<div class="order-row">
  <div class="order-num">Order: {order["order_id"]}</div>
  <div class="order-date">Order date: {order["order_date"]}</div>
</div>
<div class="sales-rec">Sales record #: {order["sales_record"]}</div>

<hr>

<table>
  <thead>
    <tr>
      <th>Item</th>
      <th class="r" style="width:72px">Quantity</th>
      <th class="r" style="width:80px">Item price</th>
      <th class="r" style="width:80px">Item total</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td>
        {order["item_title"]}<br>
        <span class="item-id">(Item ID: {order["item_id"]})</span>
      </td>
      <td class="r">{order["quantity"]}</td>
      <td class="r">{order["item_price"]}</td>
      <td class="r">{order["item_total"]}</td>
    </tr>
  </tbody>
</table>

<div class="ship-note">Buyer selected shipping service: {order["shipping_service"]}</div>

<div class="bottom">
  <div class="msg">
    <div class="msg-title">A message from {SELLER["name"]}</div>
    {SELLER["message"]}
  </div>

  <div class="qr">
    <div class="qr-txt">Check us out on Whatnot!</div>
    <img src="data:image/jpeg;base64,{qr_b64}" alt="Whatnot QR">
    <div class="qr-txt">Get a $15 credit!</div>
  </div>

  <div class="totals">
    <div class="trow"><span>Subtotal</span><span>{order["subtotal"]}</span></div>
    <div class="trow"><span>Shipping</span><span>{order["shipping_cost"]}</span></div>
    <div class="trow"><span>Sales tax (eBay collected)</span><span>{order["sales_tax"]}</span></div>
    <div class="trow grand"><span>Order total**</span><span>{order["order_total"]}</span></div>
  </div>
</div>

<div class="footnote">
  **Order total includes eBay collected tax. eBay collects and remits the tax
  to the tax authorities in accordance with applicable state law.
</div>

</body>
</html>"""


# ── Main ───────────────────────────────────────────────────────────────────────
def main():
    log.info("Starting packing slip generator")

    creds        = get_credentials()
    gc           = gspread.authorize(creds)
    drive_svc    = build("drive", "v3", credentials=creds)

    # Fetch images once per run (not per order)
    log.info("Fetching logos and QR code from Drive...")
    ebay_b64   = fetch_image_as_b64(drive_svc, EBAY_LOGO_ID)
    shratz_b64 = fetch_image_as_b64(drive_svc, SHRATZ_LOGO_ID)
    qr_b64     = fetch_image_as_b64(drive_svc, WHATNOT_QR_ID)
    log.info("Images loaded.")

    # Open sheet
    sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
    rows  = sheet.get_all_values()
    headers = rows[0]
    data_rows = rows[1:]  # skip header

    processed = 0
    skipped   = 0

    for i, row in enumerate(data_rows, start=2):  # row 2 = first data row in sheet
        order_id  = row[COL["ORDER_ID"] - 1].strip()   if len(row) >= COL["ORDER_ID"]        else ""
        generated = row[COL["SLIP_GENERATED"] - 1].strip() if len(row) >= COL["SLIP_GENERATED"] else ""

        # Skip blank rows or already-processed rows
        if not order_id:
            continue
        if generated.upper() == "YES":
            skipped += 1
            continue

        log.info(f"Processing order {order_id} (sheet row {i})")

        try:
            # Build order dict
            order = {
                "order_id":        cell(row, "ORDER_ID"),
                "order_date":      fmt_date(cell(row, "ORDER_DATE")),
                "sales_record":    cell(row, "SALES_RECORD"),
                "buyer_username":  cell(row, "BUYER_USERNAME"),
                "buyer_name":      cell(row, "BUYER_NAME"),
                "addr_line1":      cell(row, "ADDR_LINE1"),
                "addr_city":       cell(row, "ADDR_CITY"),
                "addr_state":      cell(row, "ADDR_STATE"),
                "addr_zip":        cell(row, "ADDR_ZIP"),
                "addr_country":    "United States" if cell(row, "ADDR_COUNTRY").upper() == "US" else cell(row, "ADDR_COUNTRY"),
                "buyer_phone":     fmt_phone(cell(row, "BUYER_PHONE")),
                "buyer_email":     cell(row, "BUYER_EMAIL"),
                "item_title":      cell(row, "ITEM_TITLE"),
                "item_id":         cell(row, "ITEM_ID"),
                "quantity":        cell(row, "QUANTITY"),
                "item_price":      fmt_currency(cell(row, "ITEM_PRICE")),
                "item_total":      fmt_currency(float(cell(row, "ITEM_PRICE")) * int(cell(row, "QUANTITY") or 1)),
                "shipping_service":fmt_shipping(cell(row, "SHIPPING_SERVICE")),
                "shipping_cost":   fmt_currency(cell(row, "SHIPPING_COST")),
                "subtotal":        fmt_currency(cell(row, "SUBTOTAL")),
                "sales_tax":       fmt_currency(cell(row, "SALES_TAX")),
                "order_total":     fmt_currency(cell(row, "ORDER_TOTAL")),
            }

            # Generate HTML → PDF
            html_content = build_html(order, ebay_b64, shratz_b64, qr_b64)
            pdf_bytes    = HTML(string=html_content).write_pdf()

            # Build filename: PackingSlip_<OrderID>_<Date>.pdf
            date_str  = datetime.now().strftime("%Y-%m-%d")
            filename  = f"PackingSlip_{order_id}_{date_str}.pdf"

            # Upload to Drive
            drive_url = upload_pdf(drive_svc, pdf_bytes, filename, DRIVE_FOLDER_ID)
            log.info(f"  Uploaded: {filename} → {drive_url}")

            # Update sheet: mark YES and write Drive URL
            sheet.update_cell(i, COL["SLIP_GENERATED"], "YES")
            sheet.update_cell(i, COL["DRIVE_URL"],      drive_url)

            processed += 1
            log.info(f"  Row {i} updated.")

        except Exception as e:
            log.error(f"  ERROR processing order {order_id}: {e}", exc_info=True)

    log.info(f"Done. Processed: {processed}, Already done: {skipped}")


if __name__ == "__main__":
    main()
