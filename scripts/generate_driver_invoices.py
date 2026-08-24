#!/usr/bin/env python3
"""
Generate running "Driver Invoice" PDFs for America's Trucking, LLC and keep them
up to date in Airtable's "Driver Invoices" table, every time a driver submits a
new load.

Runs standalone (no dependency on any Claude/Anthropic session or account) —
authenticates to Airtable purely via a Personal Access Token supplied through the
AIRTABLE_API_KEY environment variable (set as a GitHub Actions secret) — same
credential already used by generate_invoice_pdfs.py.

What it does, each run (intended to run frequently, e.g. every 15 minutes):
  1. Figures out the current pay week: Sunday 12:01am through Saturday 11:59pm,
     America/Chicago time. Loads submitted in that window are what a driver gets
     paid for the following week.
  2. Pulls every Load submitted (by actual submission timestamp, not the load's
     completion date) in that window that has a driver linked.
  3. Groups those loads by driver. For each driver with at least one load:
       - If nothing changed since last run (same load count as last time), does
         nothing — no rebuilt PDF, no re-notification.
       - If it's new or grown, rebuilds that driver's PDF, upserts a record on
         the "Driver Invoices" table (one per driver per week), uploads the PDF
         there, publishes a copy to the repo's GitHub Pages folder for SMS
         linking, and flags "Notify Needed" so the paired Airtable automation
         emails/texts the driver.
  4. Never sends the email/text itself — that's handled by an Airtable
     automation watching "Notify Needed", using Airtable's own connected email,
     so no separate email credential is needed here.

Safe to re-run constantly — it only touches a driver's record when their load
count for the week has actually changed.
"""
import base64
import os
import sys
import uuid
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

import requests
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import landscape, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
BASE_ID = "appoUnW9xYz6VvtHL"
TBL_LOADS = "tbl5PBqAHQgtLJdti"
TBL_DRIVERS = "tblSP11O5utOnLpOS"
TBL_DRIVER_INVOICES = "tblnZPueFMh1H6ml0"

# Loads fields
FLD_LOAD_SUBMITTED_AT = "fldReHiEUs4qbn2RP"  # "Submitted At" (CREATED_TIME formula)
FLD_LOAD_DATE = "fldZjwuJzJkOEiujc"
FLD_LOAD_DRIVERS_LINK = "fldI7NAWtT0iMcvNO"  # "Drivers" (real link)
FLD_LOAD_COMMODITY = "fldJ2FF4g1isR0HRA"
FLD_LOAD_FROM = "fld8Ol92o9Pvqj44J"
FLD_LOAD_FROM_LEGACY = "fldcgxis8zp6p86pO"
FLD_LOAD_TO = "fldfgAhc9wlORvvuU"
FLD_LOAD_TO_LEGACY = "fldJqHUdcoDhDJlIX"
FLD_LOAD_TICKET = "fldMFfApItiydDoJ6"
FLD_LOAD_WEIGHT_LOADED = "fldZyGni8JqS390SZ"
FLD_LOAD_UNLOAD_TICKET = "fld3LtVzalx9QsHIp"
FLD_LOAD_WEIGHT_UNLOADED = "fldoeV0cJISQSgHzX"
FLD_LOAD_RATE = "fldrCA8wySIPYTdBj"
FLD_LOAD_AMOUNT = "flduK7xZryS6zRuaY"
FLD_LOAD_DRIVER_PAY = "fldnA3ZAN4QJQYY5B"
FLD_LOAD_ID = "fldo4inNiEOJVUHN3"

# Drivers fields
FLD_DRV_NAME = "fldtGvGUOkFAVgs2q"
FLD_DRV_PHONE = "fldNz5RaBiu29umji"
FLD_DRV_ACTIVE = "fldyxLkLULB4VaBh8"

# Driver Invoices fields
FLD_DI_LABEL = "fld4nrBcjQi2EmaUf"
FLD_DI_DRIVER = "fldHXqVzXLcuj6byK"
FLD_DI_WEEK_START = "fldmCRod1W7hbepa1"
FLD_DI_WEEK_END = "fldMLTdJvsrljhy5K"
FLD_DI_LOADS = "fldg7r4TmrunDqR4k"
FLD_DI_LOAD_COUNT = "fldyFcqYI0sZxXS2C"
FLD_DI_PDF = "fldKQbhsfYncWyq7r"
FLD_DI_PUBLIC_LINK = "fldrUAQC9yJJlo6yy"
FLD_DI_NOTIFY = "fldvp48IogQTS2KcL"

# Driver-facing letterhead — this is America's Trucking's physical location,
# where drivers report to (different from the billing address used on
# customer invoices in generate_invoice_pdfs.py, which is intentionally left
# unchanged).
COMPANY_NAME = "America's Trucking, LLC"
COMPANY_ADDRESS = "121 S. Country Estates Rd, Liberal, KS 67901"
COMPANY_PHONE = "Phone 620-655-8268"
COMPANY_FAX = "Fax 620-626-5482"
COMPANY_EMAIL = "americastrucking@gmail.com"

TIMEZONE = ZoneInfo("America/Chicago")

# GitHub Pages hosting for the public driver-invoice PDF links. This is
# deliberately a SEPARATE, small public repo (Pages requires a public repo,
# or a paid plan, to publish from a private one) — the private repo this
# script lives in never becomes public. The workflow checks that repo out
# into DRIVER_INVOICE_DOCS_DIR and pushes changes there with its own scoped
# token (PAGES_REPO_TOKEN), never with this script's Airtable credential.
# Update GH_PAGES_BASE_URL if that repo is ever renamed or moved.
GH_PAGES_BASE_URL = "https://americastrucking.github.io/driver-invoice-files"
DOCS_DIR = os.environ.get(
    "DRIVER_INVOICE_DOCS_DIR",
    os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "driver-invoices"),
)

API_BASE = "https://api.airtable.com/v0"
CONTENT_API_BASE = "https://content.airtable.com/v0"


def airtable_headers():
    token = os.environ.get("AIRTABLE_API_KEY")
    if not token:
        print("ERROR: AIRTABLE_API_KEY environment variable is not set.", file=sys.stderr)
        sys.exit(1)
    return {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}


def airtable_list(table_id, params=None):
    params = dict(params or {})
    params["returnFieldsByFieldId"] = "true"
    records = []
    offset = None
    while True:
        if offset:
            params["offset"] = offset
        resp = requests.get(f"{API_BASE}/{BASE_ID}/{table_id}", headers=airtable_headers(), params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        records.extend(data.get("records", []))
        offset = data.get("offset")
        if not offset:
            break
    return records


def cell(record, field_id):
    return record.get("fields", {}).get(field_id)


def current_pay_week(now_central):
    """Returns (week_start, week_end) as date objects for the Sunday-Saturday
    pay week containing now_central. Sunday 12:01am through Saturday 11:59pm."""
    # Python's weekday(): Monday=0 ... Sunday=6. We want days-since-Sunday.
    days_since_sunday = (now_central.weekday() + 1) % 7
    week_start = (now_central - timedelta(days=days_since_sunday)).date()
    week_end = week_start + timedelta(days=6)
    return week_start, week_end


def resolve_location(select_field_val, legacy_text_val):
    if select_field_val and select_field_val != "Other":
        return select_field_val
    if legacy_text_val:
        return legacy_text_val
    return select_field_val or "—"


# ---------------------------------------------------------------------------
# PDF building
# ---------------------------------------------------------------------------
styles = getSampleStyleSheet()
styles.add(ParagraphStyle(name="CompanyName", fontSize=18, leading=22, fontName="Helvetica-Bold"))
styles.add(ParagraphStyle(name="SmallGray", fontSize=9, leading=12, textColor=colors.HexColor("#444444")))
styles.add(ParagraphStyle(name="DriverInvoiceTitle", fontSize=16, leading=20, fontName="Helvetica-Bold", alignment=TA_RIGHT))
styles.add(ParagraphStyle(name="RightNormal", fontSize=10, leading=13, alignment=TA_RIGHT))
styles.add(ParagraphStyle(name="BillToBody", fontSize=10, leading=14))
styles.add(ParagraphStyle(name="TicketCell", fontSize=8, leading=10, alignment=TA_CENTER, wordWrap="CJK"))


def build_driver_invoice_pdf(output_path, driver_invoice, driver, loads):
    doc = SimpleDocTemplate(
        output_path, pagesize=landscape(letter),
        leftMargin=0.5 * inch, rightMargin=0.5 * inch,
        topMargin=0.5 * inch, bottomMargin=0.5 * inch,
    )
    story = []

    header_data = [
        [Paragraph(COMPANY_NAME, styles["CompanyName"]), Paragraph("Driver Invoice", styles["DriverInvoiceTitle"])],
        [
            Paragraph(f"{COMPANY_ADDRESS}<br/>{COMPANY_PHONE} &nbsp;|&nbsp; {COMPANY_FAX}<br/>{COMPANY_EMAIL}", styles["SmallGray"]),
            Paragraph(
                f"Week of {driver_invoice['week_start']} – {driver_invoice['week_end']}",
                styles["RightNormal"],
            ),
        ],
    ]
    header_table = Table(header_data, colWidths=[6.5 * inch, 3.0 * inch])
    header_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 4))
    story.append(Table([[""]], colWidths=[9.5 * inch], rowHeights=[2],
                        style=TableStyle([("LINEBELOW", (0, 0), (-1, -1), 1.2, colors.HexColor("#222222"))])))
    story.append(Spacer(1, 14))

    driver_block = Paragraph(
        f"<b>Driver:</b><br/>{driver['name']}<br/>{driver.get('phone') or ''}",
        styles["BillToBody"],
    )
    dates_rows = [
        ["Issue Date:", Paragraph(driver_invoice["issue_date"], styles["BillToBody"])],
        ["Due Date:", Paragraph(driver_invoice["due_date"], styles["BillToBody"])],
        ["Terms:", Paragraph(driver_invoice["terms_text"], styles["BillToBody"])],
    ]
    dates_table = Table(dates_rows, colWidths=[1.3 * inch, 3.2 * inch])
    dates_table.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
    ]))
    top_block = Table([[driver_block, dates_table]], colWidths=[4.8 * inch, 4.7 * inch])
    top_block.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(top_block)
    story.append(Spacer(1, 18))

    col_headers = ["Date", "Commodity", "From", "Load\nTicket #", "Weight\nLoaded", "To",
                   "Unload\nTicket #", "Weight\nUnloaded", "Rate", "Amount", "Driver\nPay"]
    table_data = [col_headers]
    for ld in loads:
        table_data.append([
            ld["date"], ld.get("commodity") or "—", ld["from_loc"],
            Paragraph(ld["load_ticket"] or "—", styles["TicketCell"]),
            f"{ld['weight_loaded']:,}" if ld.get("weight_loaded") is not None else "—",
            ld["to_loc"],
            Paragraph(ld["unload_ticket"] or "—", styles["TicketCell"]),
            f"{ld['weight_unloaded']:,}" if ld["weight_unloaded"] is not None else "—",
            f"${ld['rate']:.2f}" if ld["rate"] is not None else "—",
            f"${ld['amount']:,.2f}" if ld["amount"] is not None else "—",
            f"${ld['driver_pay']:,.2f}" if ld["driver_pay"] is not None else "—",
        ])
    item_table = Table(
        table_data,
        colWidths=[0.75 * inch, 0.8 * inch, 0.95 * inch, 1.0 * inch, 0.65 * inch,
                   0.95 * inch, 1.0 * inch, 0.65 * inch, 0.6 * inch, 0.8 * inch, 0.8 * inch],
        repeatRows=1,
    )
    item_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F2937")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("ALIGN", (3, 0), (7, -1), "CENTER"),
        ("ALIGN", (8, 1), (10, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#999999")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F4F6")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(item_table)
    story.append(Spacer(1, 16))

    totals_rows = [
        ["Total Load Value:", f"${driver_invoice['total_load_value']:,.2f}"],
        ["Total Driver Pay Due:", f"${driver_invoice['total_driver_pay']:,.2f}"],
    ]
    totals_table = Table(totals_rows, colWidths=[1.9 * inch, 1.3 * inch])
    totals_table.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 10.5),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("LINEABOVE", (0, 1), (-1, 1), 0.75, colors.black),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    outer = Table([["", totals_table]], colWidths=[6.3 * inch, 3.2 * inch])
    story.append(outer)

    story.append(Spacer(1, 10))
    story.append(Paragraph(
        "This is a running summary of loads submitted this pay week. Final payout is "
        "processed the following week per America's Trucking's standard payroll schedule.",
        styles["SmallGray"],
    ))

    doc.build(story)
    return output_path


def upload_pdf_to_driver_invoice(record_id, pdf_path, filename):
    with open(pdf_path, "rb") as f:
        content_b64 = base64.b64encode(f.read()).decode("ascii")
    url = f"{CONTENT_API_BASE}/{BASE_ID}/{record_id}/{FLD_DI_PDF}/uploadAttachment"
    payload = {"contentType": "application/pdf", "file": content_b64, "filename": filename}
    resp = requests.post(url, headers=airtable_headers(), json=payload, timeout=60)
    resp.raise_for_status()
    return resp.json()


def upsert_driver_invoice(existing, fields):
    if existing:
        resp = requests.patch(
            f"{API_BASE}/{BASE_ID}/{TBL_DRIVER_INVOICES}/{existing['id']}",
            headers=airtable_headers(),
            json={"fields": fields},
            timeout=30,
        )
    else:
        resp = requests.post(
            f"{API_BASE}/{BASE_ID}/{TBL_DRIVER_INVOICES}",
            headers=airtable_headers(),
            json={"fields": fields},
            timeout=30,
        )
    resp.raise_for_status()
    return resp.json()


def main():
    now_central = datetime.now(TIMEZONE)
    week_start, week_end = current_pay_week(now_central)
    week_start_str = week_start.isoformat()
    week_end_str = week_end.isoformat()
    print(f"[{now_central.isoformat()}] Pay week: {week_start_str} to {week_end_str} (America/Chicago)")

    # Fetch recent loads (last 8 days is plenty to cover any current-week
    # submission regardless of when in the week this runs) and filter/group
    # precisely in Python using the real Submitted At timestamp.
    all_recent_loads = airtable_list(
        TBL_LOADS,
        params={"filterByFormula": "IS_AFTER({Submitted At}, DATEADD(TODAY(), -8, 'days'))"},
    )

    week_start_dt = datetime(week_start.year, week_start.month, week_start.day, 0, 1, tzinfo=TIMEZONE)
    week_end_dt = datetime(week_end.year, week_end.month, week_end.day, 23, 59, 59, tzinfo=TIMEZONE)

    loads_by_driver = {}
    for lr in all_recent_loads:
        submitted_raw = cell(lr, FLD_LOAD_SUBMITTED_AT)
        if not submitted_raw:
            continue
        submitted_dt = datetime.fromisoformat(submitted_raw.replace("Z", "+00:00")).astimezone(TIMEZONE)
        if not (week_start_dt <= submitted_dt <= week_end_dt):
            continue
        driver_ids = cell(lr, FLD_LOAD_DRIVERS_LINK) or []
        for did in driver_ids:
            loads_by_driver.setdefault(did, []).append((submitted_dt, lr))

    if not loads_by_driver:
        print("No loads submitted yet this pay week. Done.")
        return

    drivers = {d["id"]: d for d in airtable_list(TBL_DRIVERS)}
    existing_invoices = airtable_list(TBL_DRIVER_INVOICES)
    existing_by_key = {}
    for inv in existing_invoices:
        inv_driver_ids = cell(inv, FLD_DI_DRIVER) or []
        inv_week_start = cell(inv, FLD_DI_WEEK_START)
        for did in inv_driver_ids:
            existing_by_key[(did, inv_week_start)] = inv

    updated, unchanged = 0, 0

    for driver_id, entries in loads_by_driver.items():
        driver_rec = drivers.get(driver_id)
        if not driver_rec:
            print(f"  SKIP: driver {driver_id} not found in Drivers table.")
            continue
        driver_name = cell(driver_rec, FLD_DRV_NAME) or driver_id
        driver_phone = cell(driver_rec, FLD_DRV_PHONE)

        entries.sort(key=lambda pair: pair[0])  # chronological by submission
        load_ids_ordered = [lr["id"] for _, lr in entries]

        existing = existing_by_key.get((driver_id, week_start_str))
        prior_count = cell(existing, FLD_DI_LOAD_COUNT) if existing else None
        if existing and prior_count == len(entries):
            unchanged += 1
            continue  # nothing new for this driver since last run

        print(f"  Rebuilding Driver Invoice for {driver_name}: {len(entries)} load(s).")

        loads = []
        total_amount, total_driver_pay = 0.0, 0.0
        for _, lr in entries:
            from_val = cell(lr, FLD_LOAD_FROM)
            to_val = cell(lr, FLD_LOAD_TO)
            commodity_val = cell(lr, FLD_LOAD_COMMODITY)
            amount = cell(lr, FLD_LOAD_AMOUNT) or 0
            driver_pay = cell(lr, FLD_LOAD_DRIVER_PAY) or 0
            total_amount += amount
            total_driver_pay += driver_pay
            loads.append({
                "date": cell(lr, FLD_LOAD_DATE) or "—",
                "commodity": commodity_val,
                "from_loc": resolve_location(from_val, cell(lr, FLD_LOAD_FROM_LEGACY)),
                "to_loc": resolve_location(to_val, cell(lr, FLD_LOAD_TO_LEGACY)),
                "load_ticket": cell(lr, FLD_LOAD_TICKET),
                "unload_ticket": cell(lr, FLD_LOAD_UNLOAD_TICKET),
                "weight_loaded": cell(lr, FLD_LOAD_WEIGHT_LOADED),
                "weight_unloaded": cell(lr, FLD_LOAD_WEIGHT_UNLOADED),
                "rate": cell(lr, FLD_LOAD_RATE),
                "amount": amount,
                "driver_pay": driver_pay,
            })

        driver_invoice = {
            "week_start": week_start_str,
            "week_end": week_end_str,
            "issue_date": now_central.strftime("%-m/%-d/%Y"),
            "due_date": "Paid the following week per standard payroll schedule",
            "terms_text": "Weekly payroll",
            "total_load_value": total_amount,
            "total_driver_pay": total_driver_pay,
        }
        driver = {"name": driver_name, "phone": driver_phone}

        safe_name = driver_name.replace("/", "-").replace(" ", "_")
        local_pdf_path = f"/tmp/{safe_name}_{week_start_str}.pdf"
        pdf_filename = f"Driver_Invoice_{safe_name}_{week_start_str}.pdf"
        build_driver_invoice_pdf(local_pdf_path, driver_invoice, driver, loads)

        # Publish a copy to the repo's GitHub Pages folder under a random,
        # unguessable filename so it can be texted as a link. The repo stays
        # private; only this one folder is served publicly via Pages.
        os.makedirs(DOCS_DIR, exist_ok=True)
        public_filename = f"{uuid.uuid4().hex}.pdf"
        public_path = os.path.join(DOCS_DIR, public_filename)
        with open(local_pdf_path, "rb") as src, open(public_path, "wb") as dst:
            dst.write(src.read())
        public_url = f"{GH_PAGES_BASE_URL}/driver-invoices/{public_filename}"

        fields = {
            FLD_DI_LABEL: f"{driver_name} - Week of {week_start_str}",
            FLD_DI_DRIVER: [driver_id],
            FLD_DI_WEEK_START: week_start_str,
            FLD_DI_WEEK_END: week_end_str,
            FLD_DI_LOADS: load_ids_ordered,
            FLD_DI_LOAD_COUNT: len(entries),
            FLD_DI_PUBLIC_LINK: public_url,
            FLD_DI_NOTIFY: True,
        }
        result = upsert_driver_invoice(existing, fields)
        upload_pdf_to_driver_invoice(result["id"], local_pdf_path, pdf_filename)
        updated += 1

    print(f"\nDone. Updated: {updated}. Unchanged (skipped): {unchanged}.")


if __name__ == "__main__":
    main()
