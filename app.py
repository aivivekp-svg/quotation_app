import io
import re
from datetime import datetime, timedelta
from typing import Optional

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import (
    BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
)

APP_TITLE = "Quotation Generator – V. Purohit & Associates"

CLIENT_TYPES = [
    "PRIVATE LIMITED", "PROPRIETORSHIP", "INDIVIDUAL", "LLP", "HUF",
    "SOCIETY", "PARTNERSHIP FIRM", "FOREIGN ENTITY", "AOP/ BOI", "TRUST",
]

ACCOUNTING_PLANS = ["Monthly Accounting", "Quarterly Accounting", "Half Yearly Accounting", "Annual Accounting"]
EVENT_SERVICE = "EVENT BASED FILING"
PT_SERVICE = "PROFESSION TAX RETURNS"

FORCE_EVENT_SUBS = {
    "FILING OF TDS RETURN IN FORM 26QB",
    "FILING OF TDS RETURN IN FORM 26QC",
    "FILING OF TDS RETURN IN FORM 27Q",
}

ALIAS_DUPLICATE = {
    "LIMITED COMPANY": "PRIVATE LIMITED",
}

ACRONYMS = ["GST", "GSTR", "PTEC", "PTRC", "ADT", "ROC", "TDS", "AOC", "MGT", "26QB", "26QC", "DIR", "MSME", "KYC"]

SUBSERVICE_RENAMES = {
    "CHANGE OF ADDRESS IN GST": "GST Amendment",
    "DIR 12": "DIR 12",
    "MSME APPLICATION": "MSME Application",
    "ROC E-KYC FOR DIRECTORS": "ROC E-KYC For Directors",
}

GST_RATE_FIXED = 18
BRAND_BLUE_HEX = "#0F4C81"
BRAND_LIGHT = "#E8F0FB"


# ── Session defaults ──────────────────────────────────────────────────────────

def _ss_set(k, v):
    if k not in st.session_state:
        st.session_state[k] = v


_ss_set("editor_active", False)
_ss_set("quote_df", pd.DataFrame())
_ss_set("event_df", pd.DataFrame())
_ss_set("client_name", "")
_ss_set("client_type", "")
_ss_set("quote_no", "")
_ss_set("discount_pct", 0)
_ss_set("letterhead", False)
_ss_set("client_addr", "")
_ss_set("client_email", "")
_ss_set("client_phone", "")
_ss_set("proposal_start", "")
_ss_set("sig_bytes", None)
_ss_set("discount_reason", "")


# ── Helpers ───────────────────────────────────────────────────────────────────

def normalize_str(x):
    return (x or "").strip().upper()


def title_with_acronyms(text: str) -> str:
    if text is None:
        return ""
    t = " ".join(str(text).split()).title()
    t = re.sub(r"\bOf\b", "of", t, flags=re.IGNORECASE)
    for token in ACRONYMS:
        t = re.sub(rf"\b{re.escape(token)}\b", token, t, flags=re.IGNORECASE)
    return t


def subservice_display_override(raw_upper: str, pretty: str) -> str:
    return SUBSERVICE_RENAMES.get(raw_upper, pretty)


def service_display_override(raw_upper: str, pretty: str) -> str:
    if raw_upper == "FILING OF GSTR RETURNS":
        return "Filing of GST Returns"
    return pretty


def money_inr(n: float) -> str:
    try:
        n = float(n)
    except Exception:
        return "0"
    neg = n < 0
    n = abs(int(round(n)))
    s = str(n)
    if len(s) <= 3:
        res = s
    else:
        res = s[-3:]
        s = s[:-3]
        while len(s) > 2:
            res = s[-2:] + "," + res
            s = s[:-2]
        if s:
            res = s + "," + res
    return "-" + res if neg else res


def parse_inr(s) -> float:
    if s is None:
        return 0.0
    if isinstance(s, (int, float)):
        return float(s)
    s = str(s).strip().replace(",", "")
    if s == "":
        return 0.0
    try:
        return float(s)
    except Exception:
        return 0.0


def validity_date_str(days: int = 30) -> str:
    return (datetime.now() + timedelta(days=days)).strftime("%d-%b-%Y")


def validate_email(email: str) -> bool:
    if not email.strip():
        return True  # optional field
    return bool(re.match(r"^[\w\.-]+@[\w\.-]+\.\w{2,}$", email.strip()))


def validate_phone(phone: str) -> bool:
    if not phone.strip():
        return True  # optional field
    digits = re.sub(r"[\s\-\+\(\)]", "", phone)
    return digits.isdigit() and 7 <= len(digits) <= 15


def load_matrices():
    xl = pd.ExcelFile("matrices.xlsx")
    df_app = xl.parse("Applicability").fillna("")
    df_fees = xl.parse("Fees").fillna("")
    for df in (df_app, df_fees):
        df["Service"] = df["Service"].map(normalize_str)
        df["SubService"] = df["SubService"].map(lambda v: normalize_str(v) if pd.notna(v) else "")
        df["ClientType"] = df["ClientType"].map(normalize_str)
    df_app["Applicable"] = df_app["Applicable"].astype(str).str.upper().isin(["TRUE", "1", "YES"])
    df_fees["FeeINR"] = pd.to_numeric(df_fees["FeeINR"], errors="coerce").fillna(0.0).astype(float)
    return df_app, df_fees, "matrices.xlsx"


def split_main_vs_event(applicable: pd.DataFrame):
    ev_mask = applicable["Service"].eq(normalize_str(EVENT_SERVICE))
    return applicable.loc[~ev_mask].copy(), applicable.loc[ev_mask].copy()


def build_quotes(client_name, client_type, df_app, df_fees,
                 selected_accounting=None, selected_pt_sub=None):
    ct = normalize_str(client_type)
    applicable = (
        df_app.query("ClientType == @ct and Applicable == True")
        .loc[:, ["Service", "SubService", "ClientType"]]
        .copy()
    )
    if selected_accounting:
        sel = normalize_str(selected_accounting)
        is_acc = applicable["Service"].eq("ACCOUNTING")
        applicable = pd.concat(
            [applicable.loc[~is_acc], applicable.loc[is_acc & (applicable["SubService"] == sel)]],
            ignore_index=True,
        )
    main_app, event_app = split_main_vs_event(applicable)
    is_pt = main_app["Service"].eq(normalize_str(PT_SERVICE))
    if selected_pt_sub is not None:
        sel_pt = normalize_str(selected_pt_sub)
        main_app = pd.concat(
            [main_app.loc[~is_pt], main_app.loc[is_pt & (main_app["SubService"] == sel_pt)]],
            ignore_index=True,
        )
    force_mask = main_app["SubService"].isin(FORCE_EVENT_SUBS)
    if force_mask.any():
        event_app = pd.concat([event_app, main_app.loc[force_mask]], ignore_index=True)
        main_app = main_app.loc[~force_mask].copy()

    def _merge_and_format(df_in: pd.DataFrame) -> pd.DataFrame:
        if df_in.empty:
            return pd.DataFrame(columns=["Service", "Details", "Annual Fees (Rs.)"])
        q = df_in.merge(df_fees, on=["Service", "SubService", "ClientType"], how="left", validate="1:1")
        q["FeeINR"] = pd.to_numeric(q["FeeINR"], errors="coerce").fillna(0.0)
        svc_pretty = q["Service"].map(title_with_acronyms)
        q["Service"] = [service_display_override(raw, pretty) for raw, pretty in zip(q["Service"], svc_pretty)]
        sub_pretty = q["SubService"].map(title_with_acronyms)
        q["SubService"] = [subservice_display_override(raw, pretty) for raw, pretty in zip(q["SubService"], sub_pretty)]
        q.sort_values(["Service", "SubService"], inplace=True)
        out = (
            q.drop(columns=["ClientType"], errors="ignore")
            .rename(columns={"SubService": "Details", "FeeINR": "Annual Fees (Rs.)"})
            .loc[:, ["Service", "Details", "Annual Fees (Rs.)"]]
        )
        return out

    main_df = _merge_and_format(main_app)
    event_df = _merge_and_format(event_app)
    total = float(main_df["Annual Fees (Rs.)"].sum()) if not main_df.empty else 0.0
    return main_df, event_df, total


def compute_totals(df_selected: pd.DataFrame, discount_pct: float):
    fees_series = df_selected["Annual Fees (Rs.)"].apply(parse_inr).fillna(0.0)
    subtotal = float(fees_series.sum())
    discount_amt = round(subtotal * (discount_pct or 0) / 100.0, 2)
    taxable = max(subtotal - discount_amt, 0.0)
    gst_amt = round(taxable * GST_RATE_FIXED / 100.0, 2)
    grand = round(taxable + gst_amt, 2)
    return subtotal, discount_amt, taxable, gst_amt, grand


# ── PDF helpers ───────────────────────────────────────────────────────────────

def build_grouped_pdf_rows_compact(df: pd.DataFrame):
    rows = [["Service", "Details", "Annual Fees<br/>(Rs.)"]]
    for svc, grp in df.groupby("Service", sort=True):
        g = grp.copy()
        g["amt"] = g["Annual Fees (Rs.)"].apply(parse_inr).fillna(0.0)
        g = g.sort_values(["Details"])
        if len(g) <= 1:
            r = g.iloc[0]
            rows.append([svc, r["Details"], money_inr(r["amt"])])
        else:
            first = True
            for _, r in g.iterrows():
                rows.append([svc if first else "", r["Details"], money_inr(r["amt"])])
                first = False
    return rows


def build_event_pdf_rows(df: pd.DataFrame):
    rows = [["Details", "Fees<br/>(Rs.)"]]
    for _, row in df.iterrows():
        raw = row.get("Annual Fees (Rs.)", "")
        sraw = str(raw).strip()
        amt = parse_inr(raw)
        display_amt = "" if sraw == "" else money_inr(amt)
        detail = (str(row.get("Details", "")).strip() or str(row.get("Service", "")).strip())
        rows.append([detail, display_amt])
    return rows


def make_pdf(client_name: str, client_type: str, quote_no: str,
             df_quote: pd.DataFrame, df_event: pd.DataFrame,
             subtotal: float, discount_pct: float, discount_amt: float,
             gst_amt: float, grand: float,
             letterhead: bool = False,
             addr: Optional[str] = None,
             email: Optional[str] = None,
             phone: Optional[str] = None,
             proposal_start: Optional[str] = None,
             signature_bytes: Optional[bytes] = None,
             discount_reason: Optional[str] = None):

    import os
    from reportlab.platypus import Image
    from reportlab.lib.utils import ImageReader

    buf = io.BytesIO()
    left, right, top, bottom = 18 * mm, 18 * mm, 16 * mm, 16 * mm
    card_inset = 10 * mm
    brand_blue = colors.HexColor(BRAND_BLUE_HEX)

    def on_page(canv, doc_):
        pw, ph = A4
        canv.saveState()
        canv.setFillColor(colors.HexColor("#F7F9FC"))
        canv.rect(0, 0, pw, ph, stroke=0, fill=1)
        x, y = card_inset, card_inset
        w = pw - 2 * card_inset
        h = ph - 2 * card_inset
        canv.setFillColor(colors.white)
        canv.setStrokeColor(colors.HexColor("#E5E7EB"))
        canv.setLineWidth(1)
        canv.rect(x, y, w, h, stroke=1, fill=1)
        if letterhead:
            try:
                for n in ("logo.png", "logo.jpg", "logo.jpeg"):
                    if os.path.exists(n):
                        ir = ImageReader(n)
                        ow, oh = ir.getSize()
                        target_w = pw - 80 * mm
                        r = target_w / ow
                        target_h = oh * r
                        cx = (pw - target_w) / 2
                        cy = (ph - target_h) / 2
                        if hasattr(canv, "setFillAlpha"):
                            canv.setFillAlpha(0.07)
                        canv.drawImage(n, cx, cy, width=target_w, height=target_h,
                                       preserveAspectRatio=True, mask="auto")
                        if hasattr(canv, "setFillAlpha"):
                            canv.setFillAlpha(1)
                        break
            except Exception:
                pass
        canv.restoreState()

    def on_page_end(canv, doc_):
        canv.saveState()
        canv.setFont("Helvetica", 8)
        y_line = 20 * mm
        canv.setLineWidth(0.7)
        canv.line(18 * mm, y_line, A4[0] - 18 * mm, y_line)
        line1 = "Office No. 5, Ground Floor, Adeshwar Arcade Commercial Premises CSL, Andheri - Kurla Road,"
        line2 = "Opp. Sangam Cinema, Andheri East, Mumbai - 400093.  Email: info@vpurohit.com  |  Contact: +91-8369508539"
        canv.drawCentredString(A4[0] / 2, 12 * mm + 4, line1)
        canv.drawCentredString(A4[0] / 2, 12 * mm - 6, line2)
        canv.drawRightString(A4[0] - 18 * mm, 12 * mm - 18, f"Page {canv.getPageNumber()}")
        canv.restoreState()

    doc = BaseDocTemplate(buf, pagesize=A4, leftMargin=left, rightMargin=right,
                          topMargin=top, bottomMargin=bottom)
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
    doc.addPageTemplates(PageTemplate(id="letterhead", frames=[frame],
                                      onPage=on_page, onPageEnd=on_page_end))

    styles = getSampleStyleSheet()
    head_center = ParagraphStyle("HeadCenter", parent=styles["Normal"],
                                  alignment=TA_CENTER, fontName="Helvetica-Bold")
    right_align = ParagraphStyle("RightAlign", parent=styles["Normal"], alignment=TA_RIGHT)
    story = []

    def find_logo_path():
        for name in ("logo.png", "logo.jpg", "logo.jpeg"):
            if os.path.exists(name):
                return name
        return None

    logo_path = find_logo_path()
    if logo_path and not letterhead:
        ir = ImageReader(logo_path)
        ow, oh = ir.getSize()
        max_w, max_h = 26 * mm, 26 * mm
        r = min(max_w / ow, max_h / oh)
        story.append(Image(logo_path, width=ow * r, height=oh * r))
        story.append(Spacer(1, 4))

    story.append(Paragraph("<b>V. Purohit &amp; Associates</b>", styles["Title"]))
    story.append(Paragraph("<b>Annual Fees Proposal</b>", styles["h2"]))
    story.append(Spacer(1, 6))

    # ── Header meta table ──
    left_cells, right_cells = [], []
    left_cells.append(Paragraph(f"<b>Client Name:</b> {client_name}", styles["Normal"]))
    left_cells.append(Paragraph(f"<b>Client Entity Type:</b> {client_type}", styles["Normal"]))
    if addr and addr.strip():
        addr_html = "<br/>".join([ln.strip() for ln in addr.splitlines() if ln.strip()])
        left_cells.append(Paragraph(f"<b>Address:</b> {addr_html}", styles["Normal"]))
    if email and email.strip():
        left_cells.append(Paragraph(f"<b>Email:</b> {email.strip()}", styles["Normal"]))
    if phone and phone.strip():
        left_cells.append(Paragraph(f"<b>Phone:</b> {phone.strip()}", styles["Normal"]))

    right_cells.append(Paragraph(f"<b>Quotation No.:</b> {quote_no}", styles["Normal"]))
    right_cells.append(Paragraph(f"<b>Date:</b> {datetime.now().strftime('%d-%b-%Y')}", styles["Normal"]))
    right_cells.append(Paragraph(f"<b>Valid Until:</b> {validity_date_str(30)}", styles["Normal"]))
    if proposal_start and str(proposal_start).strip():
        right_cells.append(Paragraph(f"<b>Proposed Start:</b> {proposal_start.strip()}", styles["Normal"]))

    rows_n = max(len(left_cells), len(right_cells))
    while len(left_cells) < rows_n:
        left_cells.append(Paragraph("&nbsp;", styles["Normal"]))
    while len(right_cells) < rows_n:
        right_cells.append(Paragraph("&nbsp;", styles["Normal"]))

    meta_table = Table([[left_cells[i], right_cells[i]] for i in range(rows_n)],
                        colWidths=[110 * mm, 60 * mm])
    meta_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 8))

    # ── Main fees table ──
    table_rows = build_grouped_pdf_rows_compact(df_quote)
    table_rows[0] = [
        Paragraph("<b>Service</b>", head_center),
        Paragraph("<b>Details</b>", head_center),
        Paragraph("<b>Annual Fees</b><br/><b>(Rs.)</b>", head_center),
    ]
    col_widths = [60 * mm, 80 * mm, 30 * mm]
    table = Table(table_rows, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, 0), 10),
        ("BACKGROUND", (0, 0), (-1, 0), brand_blue),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, 0), 6),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
        ("BOX", (0, 0), (-1, 0), 0.9, brand_blue),
        ("INNERGRID", (0, 0), (-1, 0), 0.9, brand_blue),
        ("FONTSIZE", (0, 1), (-1, -1), 10),
        ("TOPPADDING", (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
        ("ALIGN", (2, 1), (2, -1), "RIGHT"),
        ("INNERGRID", (0, 1), (-1, -1), 0.3, colors.HexColor("#d9d9d9")),
        ("BOX", (0, 0), (-1, -1), 1.0, colors.black),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F7FB")]),
    ]))
    story.append(table)
    story.append(Spacer(1, 8))

    # ── Totals ──
    tot_lines = [
        ["", "Subtotal", money_inr(subtotal)],
        ["", f"Discount ({discount_pct:.0f}%)" + (f" — {discount_reason}" if discount_reason and discount_reason.strip() else ""),
         f"- {money_inr(discount_amt)}"] if discount_amt > 0 else ["", "Discount (0%)", money_inr(0)],
        ["", "Taxable Amount", money_inr(subtotal - discount_amt)],
        ["", "GST (18%)", money_inr(gst_amt)],
        ["", "Grand Total", money_inr(grand)],
    ]
    t2 = Table([["", "", ""], *tot_lines], colWidths=col_widths)
    t2.setStyle(TableStyle([
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ALIGN", (2, 0), (2, -1), "RIGHT"),
        ("FONTNAME", (1, 1), (2, 1), "Helvetica-Bold"),
        ("FONTNAME", (1, 3), (2, 3), "Helvetica-Bold"),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#E8F0FB")),
        ("TEXTCOLOR", (0, -1), (-1, -1), colors.HexColor(BRAND_BLUE_HEX)),
        ("LINEABOVE", (0, 1), (-1, 1), 0.5, colors.grey),
        ("TOPPADDING", (0, 1), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
        ("BOX", (0, 0), (-1, -1), 1.0, colors.black),
    ]))
    story.append(t2)
    story.append(Spacer(1, 8))

    # ── Notes ──
    notes = (
        "<b>Notes:</b><br/>"
        "1. All fees are exclusive of applicable taxes and out-of-pocket expenses.<br/>"
        "2. Our scope is limited strictly to the services listed above.<br/>"
        f"3. This quotation is valid until <b>{validity_date_str(30)}</b>.<br/>"
        "4. Services will commence upon execution of a formal engagement letter."
    )
    story.append(Paragraph(notes, styles["Normal"]))
    story.append(Spacer(1, 16))

    # ── Signature block ──
    sig_col_w = [95 * mm, 75 * mm]
    sig_left = [
        Paragraph("For <b>V. Purohit &amp; Associates</b>", styles["Normal"]),
        Paragraph("Chartered Accountants", styles["Normal"]),
        Spacer(1, 18),
        Paragraph("_________________________________", styles["Normal"]),
        Paragraph("<b>Authorised Signatory</b>", styles["Normal"]),
        Paragraph(datetime.now().strftime("%d-%b-%Y"), styles["Normal"]),
    ]
    sig_right = [
        Paragraph("Accepted by Client:", styles["Normal"]),
        Paragraph(f"<b>{client_name}</b>", styles["Normal"]),
        Spacer(1, 18),
        Paragraph("_________________________________", styles["Normal"]),
        Paragraph("<b>Signature &amp; Stamp</b>", styles["Normal"]),
        Paragraph("Date: ___________________", styles["Normal"]),
    ]
    sig_rows = max(len(sig_left), len(sig_right))
    while len(sig_left) < sig_rows:
        sig_left.append(Spacer(1, 4))
    while len(sig_right) < sig_rows:
        sig_right.append(Spacer(1, 4))

    sig_table = Table([[sig_left[i], sig_right[i]] for i in range(sig_rows)],
                       colWidths=sig_col_w)
    sig_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(sig_table)

    # ── Event-based page ──
    if not df_event.empty:
        story.append(PageBreak())
        story.append(Paragraph(
            "<b>Event-Based Charges</b> (applicable as and when events occur; not included in annual fees)",
            styles["Normal"]))
        story.append(Spacer(1, 4))
        ev_rows = build_event_pdf_rows(df_event)
        ev_rows[0] = [
            Paragraph("<b>Details</b>", head_center),
            Paragraph("<b>Fees</b><br/><b>(Rs.)</b>", head_center),
        ]
        ev = Table(ev_rows, colWidths=[140 * mm, 30 * mm], repeatRows=1)
        ev.setStyle(TableStyle([
            ("FONTSIZE", (0, 0), (-1, 0), 10),
            ("BACKGROUND", (0, 0), (-1, 0), brand_blue),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, 0), 6),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
            ("BOX", (0, 0), (-1, 0), 0.9, brand_blue),
            ("INNERGRID", (0, 0), (-1, 0), 0.9, brand_blue),
            ("FONTSIZE", (0, 1), (-1, -1), 10),
            ("TOPPADDING", (0, 1), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
            ("ALIGN", (1, 1), (1, -1), "RIGHT"),
            ("INNERGRID", (0, 1), (-1, -1), 0.3, colors.HexColor("#d9d9d9")),
            ("BOX", (0, 0), (-1, -1), 1.0, colors.black),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F7FB")]),
        ]))
        story.append(ev)

    doc.build(story)
    return buf.getvalue()


# ── Excel export ──────────────────────────────────────────────────────────────

def export_proposal_excel(df_main, df_event, client_name, client_type, quote_no,
                           subtotal, discount_pct, discount_amt, gst_amt, grand,
                           discount_reason=""):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    thin = Side(style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor=BRAND_BLUE_HEX.replace("#", ""))
    head_font = Font(color="FFFFFF", bold=True)
    right = Alignment(horizontal="right")
    center = Alignment(horizontal="center")

    ws = wb.active
    ws.title = "Annual Fees"
    headers = ["Service", "Details", "Annual Fees (Rs.)"]
    ws.append(headers)
    for c in range(1, 4):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font, cell.alignment, cell.border = head_fill, head_font, center, border_all

    for svc, grp in df_main.groupby("Service", sort=True):
        g = grp.sort_values(["Details"]).copy()
        first = True
        for _, r in g.iterrows():
            amt = int(round(parse_inr(r.get("Annual Fees (Rs.)", 0))))
            ws.append([svc if first else "", r.get("Details", ""), amt])
            first = False

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=3):
        row[2].number_format = '#,##,##0'
        for cell in row:
            cell.border = border_all

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 18

    start = ws.max_row + 2
    ws.cell(row=start, column=2, value="Totals").font = Font(bold=True)
    disc_label = f"Discount ({int(discount_pct)}%)"
    if discount_reason and discount_reason.strip():
        disc_label += f" — {discount_reason}"
    items = [
        ("Subtotal", subtotal),
        (disc_label, -discount_amt if discount_amt else 0),
        ("Taxable Amount", subtotal - discount_amt),
        ("GST (18%)", gst_amt),
        ("Grand Total", grand),
    ]
    for i, (lbl, amt) in enumerate(items, start=start + 1):
        ws.cell(row=i, column=2, value=lbl)
        c = ws.cell(row=i, column=3, value=int(round(amt)))
        c.number_format = '#,##,##0'
        c.alignment = right

    if not df_event.empty:
        ws2 = wb.create_sheet("Event-based charges")
        ws2.append(["Details", "Fees (Rs.)"])
        h1, h2 = ws2["A1"], ws2["B1"]
        for h in (h1, h2):
            h.fill = head_fill
            h.font = head_font
            h.alignment = center
            h.border = border_all
        for _, r in df_event.iterrows():
            detail = (str(r.get("Details", "")).strip() or str(r.get("Service", "")).strip())
            raw = str(r.get("Annual Fees (Rs.)", "")).strip()
            if raw == "":
                ws2.append([detail, None])
            else:
                amt = int(round(parse_inr(raw)))
                ws2.append([detail, amt])
        for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, min_col=1, max_col=2):
            row[1].number_format = '#,##,##0'
            for cell in row:
                cell.border = border_all
        ws2.column_dimensions["A"].width, ws2.column_dimensions["B"].width = 60, 18

    ws3 = wb.create_sheet("Cover")
    meta = [
        ["Client Name", client_name],
        ["Client Entity Type", client_type],
        ["Quotation No.", quote_no],
        ["Date", datetime.now().strftime("%d-%b-%Y")],
        ["Valid Until", validity_date_str(30)],
    ]
    for r in meta:
        ws3.append(r)
    ws3.column_dimensions["A"].width = 22
    ws3.column_dimensions["B"].width = 50

    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def build_status(df_app, df_fees):
    active = df_app[df_app["Applicable"] == True].copy()
    counts = (active.groupby("ClientType").size()
              .reindex([normalize_str(x) for x in CLIENT_TYPES], fill_value=0)
              .reset_index(name="Applicable services"))
    merged = active.merge(df_fees, on=["Service", "SubService", "ClientType"], how="left")
    missing_mask = (merged["FeeINR"].isna() |
                    (pd.to_numeric(merged["FeeINR"], errors="coerce").fillna(0.0) <= 0))
    miss = (merged[missing_mask].groupby("ClientType").size()
            .reindex([normalize_str(x) for x in CLIENT_TYPES], fill_value=0)
            .reset_index(name="Missing/Zero fees"))
    status = counts.merge(miss, on="ClientType")
    status["ClientType"] = status["ClientType"].str.title()
    return status


# ── Page config & CSS ─────────────────────────────────────────────────────────

st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="centered")

st.markdown(f"""
<style>
  :root {{ color-scheme: only light; }}
  html, body, .stApp, .block-container {{
      background: #FFFFFF !important; color: #111111 !important;
  }}
  h1, h2, h3, h4, h5, h6 {{ color: #0E0E0E !important; }}
  div[data-testid="stDataFrame"] * {{ color: #111111 !important; }}

  /* Buttons */
  .stButton > button, .stDownloadButton > button {{
      background: {BRAND_BLUE_HEX} !important;
      color: #fff !important;
      border: 0;
      border-radius: 6px;
      font-weight: 600;
  }}
  .stButton > button:hover, .stDownloadButton > button:hover {{
      background: #0a3a6b !important;
  }}

  /* Hide heading anchor links */
  h1 > a, h2 > a, h3 > a, h4 > a {{ display: none !important; }}

  /* Step badge */
  .step-badge {{
      display: inline-block;
      background: {BRAND_BLUE_HEX};
      color: white;
      font-weight: 700;
      font-size: 0.78rem;
      border-radius: 50%;
      width: 22px; height: 22px;
      text-align: center;
      line-height: 22px;
      margin-right: 6px;
  }}
  .step-header {{
      font-size: 1.05rem;
      font-weight: 700;
      margin: 14px 0 4px 0;
      display: flex;
      align-items: center;
  }}

  /* Section divider */
  .section-divider {{
      border: none;
      border-top: 2px solid {BRAND_BLUE_HEX};
      margin: 18px 0 10px 0;
      opacity: 0.15;
  }}

  /* Info box */
  .info-box {{
      background: {BRAND_LIGHT};
      border-left: 4px solid {BRAND_BLUE_HEX};
      padding: 10px 14px;
      border-radius: 4px;
      font-size: 0.9rem;
      margin: 8px 0 12px 0;
  }}

  /* Totals card */
  .totals-card {{
      background: {BRAND_LIGHT};
      border: 1px solid #c9d9f0;
      border-radius: 8px;
      padding: 16px 20px;
      margin: 12px 0;
  }}
  .totals-card .grand {{
      font-size: 1.2rem;
      font-weight: 800;
      color: {BRAND_BLUE_HEX};
  }}
</style>
""", unsafe_allow_html=True)


# ── Load matrices ─────────────────────────────────────────────────────────────

try:
    df_app, df_fees, source = load_matrices()
except Exception as e:
    st.error(f"⚠️ Error loading matrices.xlsx: {e}")
    st.stop()

# Alias duplication & rebuild CLIENT_TYPES
for new_type, base_type in ALIAS_DUPLICATE.items():
    base = normalize_str(base_type)
    new = normalize_str(new_type)
    rows_app = df_app[df_app["ClientType"] == base].copy()
    rows_fees = df_fees[df_fees["ClientType"] == base].copy()
    if not rows_app.empty:
        rows_app["ClientType"] = new
        df_app = pd.concat([df_app, rows_app], ignore_index=True)
    if not rows_fees.empty:
        rows_fees["ClientType"] = new
        df_fees = pd.concat([df_fees, rows_fees], ignore_index=True)

CLIENT_TYPES = sorted(df_app["ClientType"].dropna().unique().tolist())


# ── Sidebar ───────────────────────────────────────────────────────────────────

with st.sidebar:
    st.image("logo.png") if __import__("os").path.exists("logo.png") else None
    st.markdown("### V. Purohit & Associates")
    st.caption("Annual Fees Quotation Tool")
    st.divider()

    st.markdown("#### ⚙️ Proposal Options")

    st.session_state["discount_pct"] = st.number_input(
        "Discount %", 0, 100, int(st.session_state["discount_pct"]), 1,
        help="Applied to subtotal before GST"
    )

    if st.session_state["discount_pct"] > 0:
        st.session_state["discount_reason"] = st.text_input(
            "Reason for discount",
            st.session_state.get("discount_reason", ""),
            placeholder="e.g., Introductory offer"
        )

    st.session_state["letterhead"] = st.checkbox(
        "Letterhead mode (watermark logo)",
        value=st.session_state["letterhead"]
    )

    sig_up = st.file_uploader("Signature / Stamp image (optional)", type=["png", "jpg", "jpeg"])
    if sig_up is not None:
        st.session_state["sig_bytes"] = sig_up.read()

    st.divider()

    with st.expander("📊 Data Status", expanded=False):
        service_defs = len(df_app[["Service", "SubService"]].drop_duplicates())
        st.write(f"**Source:** {source}")
        st.write(f"**Service definitions:** {service_defs}")
        status_df = build_status(df_app, df_fees)
        all_ok = (status_df["Missing/Zero fees"] == 0).all()
        if all_ok:
            st.success("All fees are mapped correctly.")
            st.dataframe(status_df.drop(columns=["Missing/Zero fees"]), use_container_width=True)
        else:
            st.warning("Some fees are missing or zero.")
            st.dataframe(status_df, use_container_width=True)

    st.divider()
    st.markdown(
        "<small>**How to use:**<br/>"
        "① Fill client details<br/>"
        "② Select accounting plan & PT type<br/>"
        "③ Click Generate Table<br/>"
        "④ Edit fees if needed<br/>"
        "⑤ Download PDF or Excel</small>",
        unsafe_allow_html=True
    )


# ── Main header ───────────────────────────────────────────────────────────────

st.markdown(f"""
<div style="background:{BRAND_BLUE_HEX};padding:18px 22px;border-radius:8px;margin-bottom:18px;">
  <h2 style="color:white;margin:0;font-size:1.35rem;">📄 Quotation Generator</h2>
  <p style="color:#cce0ff;margin:4px 0 0 0;font-size:0.9rem;">V. Purohit &amp; Associates — Chartered Accountants</p>
</div>
""", unsafe_allow_html=True)

# Step indicator
if not st.session_state["editor_active"]:
    active_step = 1
else:
    active_step = 2

steps = ["Client Details", "Review & Edit Fees", "Download"]
cols_steps = st.columns(len(steps))
for i, (col, step) in enumerate(zip(cols_steps, steps), start=1):
    with col:
        bg = BRAND_BLUE_HEX if i == active_step else ("#28a745" if i < active_step else "#dee2e6")
        text_col = "white" if i <= active_step else "#666"
        st.markdown(
            f'<div style="text-align:center;background:{bg};color:{text_col};'
            f'padding:7px 4px;border-radius:6px;font-size:0.82rem;font-weight:600;">'
            f'{"✓ " if i < active_step else f"{i}. "}{step}</div>',
            unsafe_allow_html=True
        )

st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)


# ── Input Form ────────────────────────────────────────────────────────────────

with st.form("quote_form", clear_on_submit=False):

    st.markdown("#### 👤 Client Information")

    col_a, col_b = st.columns([3, 2])
    with col_a:
        client_name = st.text_input(
            "Client Name *",
            st.session_state.get("client_name", ""),
            placeholder="Enter full legal name of client"
        )
    with col_b:
        ct_index = CLIENT_TYPES.index(st.session_state.get("client_type", CLIENT_TYPES[0])) \
            if st.session_state.get("client_type", "") in CLIENT_TYPES else 0
        client_type = st.selectbox("Client Type *", CLIENT_TYPES, index=ct_index)

    addr = st.text_area(
        "Address",
        st.session_state.get("client_addr", ""),
        placeholder="Street, Area\nCity, State, PIN",
        height=80
    )

    col_c, col_d = st.columns(2)
    with col_c:
        email = st.text_input(
            "Email",
            st.session_state.get("client_email", ""),
            placeholder="client@example.com"
        )
    with col_d:
        phone = st.text_input(
            "Phone",
            st.session_state.get("client_phone", ""),
            placeholder="+91 98765 43210"
        )

    proposal_start = st.text_input(
        "Proposed Engagement Start",
        st.session_state.get("proposal_start", ""),
        placeholder="e.g., 1st April 2025 / FY 2025-26"
    )

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
    st.markdown("#### 📋 Service Selection")

    ct_norm = normalize_str(client_type)
    app_ct = df_app[(df_app["ClientType"] == ct_norm) & (df_app["Applicable"] == True)]

    col_e, col_f = st.columns(2)
    with col_e:
        st.markdown("**Accounting Plan** — select one")
        selected_accounting = st.radio(
            "Accounting Plan",
            ACCOUNTING_PLANS,
            index=3,
            horizontal=False,
            label_visibility="collapsed"
        )
    with col_f:
        st.markdown("**Profession Tax Returns** — select one")
        pt_options = app_ct.loc[app_ct["Service"] == normalize_str(PT_SERVICE), "SubService"].dropna().unique().tolist()
        pt_options_tc = sorted([title_with_acronyms(s) for s in pt_options if s])
        selected_pt_tc = st.radio(
            "PT Returns",
            pt_options_tc if pt_options_tc else ["(Not applicable)"],
            index=0,
            horizontal=False,
            label_visibility="collapsed"
        )
        if selected_pt_tc == "(Not applicable)":
            selected_pt_tc = None

    st.markdown("<br/>", unsafe_allow_html=True)
    submit = st.form_submit_button("⚡ Generate Quotation Table", use_container_width=True)


# ── Generate on submit ────────────────────────────────────────────────────────

if submit:
    errors = []
    if not client_name.strip():
        errors.append("Client Name is required.")
    if not validate_email(email):
        errors.append(f"Email address '{email}' does not appear valid.")
    if not validate_phone(phone):
        errors.append(f"Phone number '{phone}' does not appear valid.")

    if errors:
        for err in errors:
            st.error(f"⚠️ {err}")
    else:
        st.session_state["quote_no"] = datetime.now().strftime("QTN-%Y%m%d-%H%M%S")
        st.session_state["proposal_start"] = proposal_start
        main_df, event_df, _ = build_quotes(
            client_name, client_type, df_app, df_fees,
            selected_accounting=selected_accounting,
            selected_pt_sub=selected_pt_tc,
        )

        # Add Consulting Charges row to event
        consulting_row = pd.DataFrame([{
            "Service": "Consulting Charges",
            "Details": "Consulting Charges",
            "Annual Fees (Rs.)": ""
        }])
        if event_df.empty:
            event_df = consulting_row.copy()
        else:
            if not event_df["Details"].astype(str).str.strip().str.casefold().eq("consulting charges").any():
                event_df = pd.concat([event_df, consulting_row], ignore_index=True)

        if main_df.empty and event_df.empty:
            st.warning("No applicable services found for the selected Client Type.")
        else:
            main_df = main_df.copy()
            main_df["Annual Fees (Rs.)"] = main_df["Annual Fees (Rs.)"].map(lambda x: money_inr(float(x)))
            main_df["Include"] = True
            main_df["MoveToEvent"] = False
            st.session_state["quote_df"] = main_df

            ev = event_df.copy()
            if not ev.empty:
                def _fmt_blank_preserve(v):
                    s = str(v).strip()
                    if s == "" or s.lower() == "nan":
                        return ""
                    try:
                        return money_inr(float(str(v).replace(",", "")))
                    except Exception:
                        return ""
                ev["Annual Fees (Rs.)"] = ev["Annual Fees (Rs.)"].apply(_fmt_blank_preserve)
                ev["MoveToMain"] = False
            st.session_state["event_df"] = ev
            st.session_state["client_name"] = client_name
            st.session_state["client_type"] = client_type
            st.session_state["client_addr"] = addr
            st.session_state["client_email"] = email
            st.session_state["client_phone"] = phone
            st.session_state["editor_active"] = True


# ── Editors / Totals / Exports ────────────────────────────────────────────────

if st.session_state["editor_active"] and (
    not st.session_state["quote_df"].empty or not st.session_state["event_df"].empty
):
    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    col_hdr, col_restart = st.columns([5, 1])
    with col_hdr:
        st.markdown(f"#### 📝 Review Fees — {st.session_state['client_name']} ({st.session_state['client_type'].title()})")
    with col_restart:
        if st.button("↩ Start Over"):
            st.session_state["editor_active"] = False
            st.session_state["quote_df"] = pd.DataFrame()
            st.session_state["event_df"] = pd.DataFrame()
            st.rerun()

    st.markdown(
        '<div class="info-box">✏️ You can edit fees directly in the table below. '
        'Use Indian format (e.g., <b>1,00,000</b>). '
        'Uncheck <b>Include</b> to exclude a service from the proposal. '
        'Use <b>MoveToEvent</b> to shift a service to the event-based section.</div>',
        unsafe_allow_html=True
    )

    # Main editor
    if not st.session_state["quote_df"].empty:
        with st.form("edit_main"):
            edited = st.data_editor(
                st.session_state["quote_df"],
                use_container_width=True,
                disabled=["Service", "Details"],
                column_order=["Include", "MoveToEvent", "Service", "Details", "Annual Fees (Rs.)"],
                column_config={
                    "Include": st.column_config.CheckboxColumn(
                        "Include",
                        help="Uncheck to remove from proposal/PDF."
                    ),
                    "MoveToEvent": st.column_config.CheckboxColumn(
                        "→ Event",
                        help="Tick and apply to move to Event-based section."
                    ),
                    "Annual Fees (Rs.)": st.column_config.TextColumn(
                        "Annual Fees (Rs.)",
                        help="Indian format, digits and commas only.",
                        validate=r"^\s*[\d,]*\s*$",
                    ),
                },
                num_rows="fixed",
                key="quote_editor",
                hide_index=True,
                height=420,
            )
            col_m1, col_m2 = st.columns(2)
            apply_edits = col_m1.form_submit_button("✅ Apply Edits", use_container_width=True)
            apply_and_move = col_m2.form_submit_button("✅ Apply & Move Selected to Event", use_container_width=True)

            if apply_edits or apply_and_move:
                edited = edited.copy()
                edited["Annual Fees (Rs.)"] = edited["Annual Fees (Rs.)"].apply(
                    lambda x: money_inr(parse_inr(x))
                )
                st.session_state["quote_df"] = edited

                if apply_and_move:
                    move_rows = edited[edited["MoveToEvent"] == True].copy()
                    if not move_rows.empty:
                        addon = move_rows[["Service", "Details", "Annual Fees (Rs.)"]].copy()
                        addon["Details"] = addon["Details"].apply(
                            lambda x: x.strip() if isinstance(x, str) else ""
                        )
                        addon.loc[addon["Details"] == "", "Details"] = addon["Service"]
                        ev_now = st.session_state["event_df"].copy()
                        if ev_now.empty:
                            ev_now = pd.DataFrame(columns=["Service", "Details", "Annual Fees (Rs.)", "MoveToMain"])
                        addon["MoveToMain"] = False
                        st.session_state["event_df"] = pd.concat([ev_now, addon], ignore_index=True)
                        kept = edited.loc[edited["MoveToEvent"] != True].copy()
                        kept["MoveToEvent"] = False
                        st.session_state["quote_df"] = kept
                        st.success(f"Moved {len(addon)} row(s) to Event-based section.")

    qdf = st.session_state["quote_df"]
    filtered = qdf[qdf["Include"] == True].copy() if "Include" in qdf.columns else qdf.copy()
    for col in ["Include", "MoveToEvent"]:
        if col in filtered.columns:
            filtered.drop(columns=[col], inplace=True)

    # Event editor
    event_df = st.session_state["event_df"].copy()
    if not event_df.empty:
        st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
        st.markdown("#### 📌 Event-Based Charges")
        st.caption("These are charged as and when events occur — not included in annual totals.")

        if "MoveToMain" not in event_df.columns:
            event_df["MoveToMain"] = False

        with st.form("event_form"):
            event_edited = st.data_editor(
                event_df,
                use_container_width=True,
                disabled=["Service", "Details"],
                column_order=["MoveToMain", "Service", "Details", "Annual Fees (Rs.)"],
                column_config={
                    "MoveToMain": st.column_config.CheckboxColumn(
                        "→ Main",
                        help="Tick and apply to move back to Main table."
                    ),
                    "Annual Fees (Rs.)": st.column_config.TextColumn(
                        "Fees (Rs.)",
                        help="Leave blank if not decided. Indian format only.",
                        validate=r"^\s*[\d,]*\s*$",
                    ),
                },
                num_rows="fixed",
                key="event_editor",
                hide_index=True,
                height=320,
            )
            ev_apply = st.form_submit_button("✅ Apply Event Edits / Move to Main", use_container_width=True)

            if ev_apply:
                event_edited = event_edited.copy()

                def _fmt_blank_preserve2(v):
                    s = str(v).strip()
                    if s == "" or s.lower() == "nan":
                        return ""
                    try:
                        return money_inr(float(str(v).replace(",", "")))
                    except Exception:
                        return ""

                event_edited["Annual Fees (Rs.)"] = event_edited["Annual Fees (Rs.)"].apply(_fmt_blank_preserve2)
                to_main = event_edited[event_edited["MoveToMain"] == True].copy()
                keep_ev = event_edited[event_edited["MoveToMain"] != True].copy()

                if not to_main.empty:
                    add_main = to_main[["Service", "Details", "Annual Fees (Rs.)"]].copy()
                    add_main["Details"] = add_main["Details"].apply(
                        lambda x: x.strip() if isinstance(x, str) else ""
                    )
                    add_main.loc[add_main["Details"] == "", "Details"] = add_main["Service"]
                    main_now = st.session_state["quote_df"].copy()
                    if main_now.empty:
                        main_now = pd.DataFrame(
                            columns=["Include", "MoveToEvent", "Service", "Details", "Annual Fees (Rs.)"]
                        )
                    add_main["Include"] = True
                    add_main["MoveToEvent"] = False
                    cols = ["Include", "MoveToEvent", "Service", "Details", "Annual Fees (Rs.)"]
                    main_now = pd.concat([main_now, add_main[cols]], ignore_index=True)
                    st.session_state["quote_df"] = main_now
                    st.success(f"Moved {len(add_main)} row(s) to Main.")

                if "MoveToMain" in keep_ev.columns:
                    keep_ev["MoveToMain"] = False
                st.session_state["event_df"] = keep_ev

    # ── Totals card ──────────────────────────────────────────────────────────

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)

    calc_df = filtered.copy()
    calc_df["Annual Fees (Rs.)"] = calc_df["Annual Fees (Rs.)"].apply(parse_inr).fillna(0.0)
    subtotal, discount_amt, taxable, gst_amt, grand = compute_totals(
        calc_df, st.session_state["discount_pct"]
    )

    disc_pct = st.session_state["discount_pct"]
    disc_reason = st.session_state.get("discount_reason", "")

    st.markdown(f"""
    <div class="totals-card">
      <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
        <span>Subtotal</span><span><b>Rs. {money_inr(subtotal)}</b></span>
      </div>
      <div style="display:flex;justify-content:space-between;margin-bottom:4px;color:#666;">
        <span>Discount ({disc_pct}%){" — " + disc_reason if disc_reason else ""}</span>
        <span>- Rs. {money_inr(discount_amt)}</span>
      </div>
      <div style="display:flex;justify-content:space-between;margin-bottom:4px;">
        <span>Taxable Amount</span><span><b>Rs. {money_inr(taxable)}</b></span>
      </div>
      <div style="display:flex;justify-content:space-between;margin-bottom:8px;color:#666;">
        <span>GST @ 18%</span><span>Rs. {money_inr(gst_amt)}</span>
      </div>
      <hr style="border:none;border-top:1.5px solid {BRAND_BLUE_HEX};opacity:0.3;margin:8px 0;">
      <div style="display:flex;justify-content:space-between;" class="grand">
        <span>Grand Total</span><span>Rs. {money_inr(grand)}</span>
      </div>
      <div style="color:#888;font-size:0.8rem;margin-top:6px;">
        Quotation valid until: <b>{validity_date_str(30)}</b>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Download buttons ──────────────────────────────────────────────────────

    st.markdown("<hr class='section-divider'>", unsafe_allow_html=True)
    st.markdown("#### ⬇️ Download Proposal")

    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        if not filtered.empty or not st.session_state["event_df"].empty:
            try:
                excel_bytes = export_proposal_excel(
                    filtered, st.session_state["event_df"],
                    st.session_state["client_name"], st.session_state["client_type"],
                    st.session_state.get("quote_no", ""),
                    subtotal, float(st.session_state["discount_pct"]),
                    discount_amt, gst_amt, grand,
                    discount_reason=disc_reason
                )
                st.download_button(
                    "📊 Download Excel",
                    data=excel_bytes,
                    file_name=f"Fees_Proposal_{st.session_state['client_name'].replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_xlsx_full",
                    use_container_width=True
                )
            except Exception:
                st.warning("Excel export requires 'openpyxl' on the server.")

    with col_dl2:
        pdf_bytes = make_pdf(
            st.session_state["client_name"], st.session_state["client_type"],
            st.session_state.get("quote_no", datetime.now().strftime("QTN-%Y%m%d-%H%M%S")),
            filtered, st.session_state["event_df"],
            subtotal, float(st.session_state["discount_pct"]), discount_amt, gst_amt, grand,
            letterhead=st.session_state["letterhead"],
            addr=st.session_state.get("client_addr", ""),
            email=st.session_state.get("client_email", ""),
            phone=st.session_state.get("client_phone", ""),
            proposal_start=st.session_state.get("proposal_start", ""),
            discount_reason=disc_reason
        )
        st.download_button(
            "📄 Download PDF",
            data=pdf_bytes,
            file_name=f"Fees_Proposal_{st.session_state['client_name'].replace(' ', '_')}.pdf",
            mime="application/pdf",
            key="dl_pdf",
            use_container_width=True
        )

    # WhatsApp share link
    if grand > 0:
        client_n = st.session_state['client_name']
        wa_text = (
            f"Dear {client_n},%0A%0A"
            f"Please find our annual fees proposal for {st.session_state['client_type'].title()} entity.%0A%0A"
            f"Grand Total (incl. GST): Rs. {money_inr(grand)}%0A"
            f"Valid Until: {validity_date_str(30)}%0A%0A"
            f"V. Purohit %26 Associates"
        )
        st.markdown(
            f'<a href="https://wa.me/?text={wa_text}" target="_blank" '
            f'style="display:inline-block;background:#25D366;color:white;'
            f'padding:8px 18px;border-radius:6px;text-decoration:none;'
            f'font-weight:600;font-size:0.9rem;margin-top:8px;">'
            f'💬 Share via WhatsApp</a>',
            unsafe_allow_html=True
        )
