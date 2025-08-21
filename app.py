import io
import re
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
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

# Force these to Event-based
FORCE_EVENT_SUBS = {
    "FILING OF TDS RETURN IN FORM 26QB",
    "FILING OF TDS RETURN IN FORM 26QC",
    "FILING OF TDS RETURN IN FORM 27Q",
}

# Text casing/renames
ACRONYMS = ["GST", "GSTR", "PTEC", "PTRC", "ADT", "ROC", "TDS", "AOC", "MGT", "26QB", "26QC", "DIR", "MSME", "KYC"]
SUBSERVICE_RENAMES = {
    "CHANGE OF ADDRESS IN GST": "GST Amendment",
    "DIR 12": "DIR 12",
    "MSME APPLICATION": "MSME Application",
    "ROC E-KYC FOR DIRECTORS": "ROC E-KYC For Directors",
}

GST_RATE_FIXED = 18
BRAND_BLUE_HEX = "#0F4C81"

# --------- Session defaults ---------
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

# --------- Helpers ---------
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
    if raw_upper in SUBSERVICE_RENAMES:
        return SUBSERVICE_RENAMES[raw_upper]
    return pretty

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
    # Accounting → keep one plan
    if selected_accounting:
        sel = normalize_str(selected_accounting)
        is_acc = applicable["Service"].eq("ACCOUNTING")
        applicable = pd.concat(
            [applicable.loc[~is_acc], applicable.loc[is_acc & (applicable["SubService"] == sel)]],
            ignore_index=True,
        )

    # Split Event-based
    main_app, event_app = split_main_vs_event(applicable)

    # Profession Tax Returns → choose one (in main)
    is_pt = main_app["Service"].eq(normalize_str(PT_SERVICE))
    if selected_pt_sub is not None:
        sel_pt = normalize_str(selected_pt_sub)
        main_app = pd.concat(
            [main_app.loc[~is_pt], main_app.loc[is_pt & (main_app["SubService"] == sel_pt)]],
            ignore_index=True,
        )

    # Force 26QB/26QC/27Q to Event-based
    force_mask = main_app["SubService"].isin(FORCE_EVENT_SUBS)
    if force_mask.any():
        event_app = pd.concat([event_app, main_app.loc[force_mask]], ignore_index=True)
        main_app = main_app.loc[~force_mask].copy()

    # Merge + beautify
    def _merge_and_format(df_in: pd.DataFrame) -> pd.DataFrame:
        if df_in.empty:
            return pd.DataFrame(columns=["Service", "Details", "Annual Fees (Rs.)"])
        q = df_in.merge(df_fees, on=["Service", "SubService", "ClientType"], how="left", validate="1:1")
        q["FeeINR"] = pd.to_numeric(q["FeeINR"], errors="coerce").fillna(0.0)
        service_key_upper = q["Service"].copy()
        svc_pretty = service_key_upper.map(title_with_acronyms)
        q["Service"] = [service_display_override(raw, pretty) for raw, pretty in zip(service_key_upper.tolist(), svc_pretty.tolist())]
        raw_sub_upper = q["SubService"].copy()
        sub_pretty = raw_sub_upper.map(title_with_acronyms)
        q["SubService"] = [subservice_display_override(raw, pretty) for raw, pretty in zip(raw_sub_upper.tolist(), sub_pretty.tolist())]
        q.sort_values(["Service", "SubService"], inplace=True)
        return (
            q.drop(columns=["ClientType"], errors="ignore")
            .rename(columns={"SubService": "Details", "FeeINR": "Annual Fees (Rs.)"})
            .loc[:, ["Service", "Details", "Annual Fees (Rs.)"]]
        )

    main_df = _merge_and_format(main_app)
    event_df = _merge_and_format(event_app)
    total = float(main_df["Annual Fees (Rs.)"].sum()) if not main_df.empty else 0.0
    return main_df, event_df, total

def compute_totals(df_selected: pd.DataFrame, discount_pct: float):
    fees_series = pd.to_numeric(df_selected["Annual Fees (Rs.)"], errors="coerce").fillna(0.0)
    subtotal = float(fees_series.sum())
    discount_amt = round(subtotal * (discount_pct or 0) / 100.0, 2)
    taxable = max(subtotal - discount_amt, 0.0)
    gst_amt = round(taxable * GST_RATE_FIXED / 100.0, 2)
    grand = round(taxable + gst_amt, 2)
    return subtotal, discount_amt, taxable, gst_amt, grand

# --------- PDF helpers ---------
def build_grouped_pdf_rows_compact(df: pd.DataFrame):
    rows = [["Service", "Details", "Annual Fees<br/>(Rs.)"]]
    for svc, grp in df.groupby("Service", sort=True):
        g = grp.copy()
        g["amt"] = pd.to_numeric(g["Annual Fees (Rs.)"], errors="coerce").fillna(0.0)
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
        amt = pd.to_numeric(row["Annual Fees (Rs.)"], errors="coerce")
        amt = 0.0 if pd.isna(amt) else float(amt)
        detail = (str(row.get("Details", "")).strip() or str(row.get("Service", "")).strip())
        rows.append([detail, money_inr(amt)])
    return rows

def make_pdf(client_name: str, client_type: str, quote_no: str,
             df_quote: pd.DataFrame, df_event: pd.DataFrame,
             subtotal: float, discount_pct: float, discount_amt: float, gst_amt: float, grand: float,
             letterhead: bool = False, addr: Optional[str] = None, email: Optional[str] = None,
             phone: Optional[str] = None, proposal_start: Optional[str] = None,
             signature_bytes: Optional[bytes] = None):
    import os
    from reportlab.platypus import Image
    from reportlab.lib.utils import ImageReader

    buf = io.BytesIO()

    left, right, top, bottom = 18*mm, 18*mm, 16*mm, 16*mm
    card_inset = 10*mm
    brand_blue = colors.HexColor(BRAND_BLUE_HEX)

    def on_page(canv, doc_):
        pw, ph = A4
        canv.saveState()
        canv.setFillColor(colors.HexColor("#F7F9FC"))
        canv.rect(0, 0, pw, ph, stroke=0, fill=1)
        x = card_inset; y = card_inset
        w = pw - 2*card_inset; h = ph - 2*card_inset
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
                        target_w = pw - 80*mm
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
        y_line = 20*mm
        canv.setLineWidth(0.7)
        canv.line(18*mm, y_line, A4[0] - 18*mm, y_line)
        line1 = "Office No. 5, Ground Floor, Adeshwar Arcade Commercial Premises CSL, Andheri - Kurla Road,"
        line2 = "Opp. Sangam Cinema, Andheri East, Mumbai - 400093.  Email: info@vpurohit.com, Contact: +91-8369508539"
        canv.drawCentredString(A4[0] / 2, 12*mm + 4, line1)
        canv.drawCentredString(A4[0] / 2, 12*mm - 6, line2)
        canv.drawRightString(A4[0] - 18*mm, 12*mm - 18, f"Page {canv.getPageNumber()}")
        canv.restoreState()

    doc = BaseDocTemplate(
        buf, pagesize=A4, leftMargin=left, rightMargin=right, topMargin=top, bottomMargin=bottom
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
    doc.addPageTemplates(PageTemplate(id="letterhead", frames=[frame], onPage=on_page, onPageEnd=on_page_end))

    styles = getSampleStyleSheet()
    head_center = ParagraphStyle("HeadCenter", parent=styles["Normal"], alignment=TA_CENTER, fontName="Helvetica-Bold")
    story = []

    def find_logo_path():
        import os
        for name in ("logo.png", "logo.jpg", "logo.jpeg"):
            if os.path.exists(name):
                return name
        return None

    logo_path = find_logo_path()
    if logo_path and not letterhead:
        ir = ImageReader(logo_path); ow, oh = ir.getSize()
        max_w, max_h = 26*mm, 26*mm
        r = min(max_w/ow, max_h/oh)
        story.append(Image(logo_path, width=ow*r, height=oh*r))
        story.append(Spacer(1, 4))

    story.append(Paragraph("<b>V. Purohit & Associates</b>", styles["Title"]))
    story.append(Paragraph("<b>Annual Fees Proposal</b>", styles["h2"]))
    story.append(Spacer(1, 6))

    # Header meta
    left_cells = []
    right_cells = []
    left_cells.append(Paragraph(f"<b>Client Name:</b> {client_name}", styles["Normal"]))
    left_cells.append(Paragraph(f"<b>Client Entity Type:</b> {client_type}", styles["Normal"]))
    addr_html = ""
    if addr and addr.strip():
        addr_html = "<br/>".join([ln.strip() for ln in addr.splitlines() if ln.strip()])
        left_cells.append(Paragraph(f"<b>Address:</b> {addr_html}", styles["Normal"]))
    if email and email.strip():
        left_cells.append(Paragraph(f"<b>Email:</b> {email.strip()}", styles["Normal"]))
    if phone and phone.strip():
        left_cells.append(Paragraph(f"<b>Phone:</b> {phone.strip()}", styles["Normal"]))
    right_cells.append(Paragraph(f"<b>Quotation No.:</b> {quote_no}", styles["Normal"]))
    right_cells.append(Paragraph(f"<b>Date:</b> {datetime.now().strftime('%d-%b-%Y')}", styles["Normal"]))
    if proposal_start and str(proposal_start).strip():
        right_cells.append(Paragraph(f"<b>Proposed Start:</b> {proposal_start.strip()}", styles["Normal"]))
    rows_n = max(len(left_cells), len(right_cells))
    while len(left_cells) < rows_n: left_cells.append(Paragraph("&nbsp;", styles["Normal"]))
    while len(right_cells) < rows_n: right_cells.append(Paragraph("&nbsp;", styles["Normal"]))
    meta_table = Table([[left_cells[i], right_cells[i]] for i in range(rows_n)], colWidths=[110*mm, 60*mm])
    meta_table.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "TOP"),
        ("LEFTPADDING", (0,0), (-1,-1), 0),
        ("RIGHTPADDING", (0,0), (-1,-1), 0),
        ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("TOPPADDING", (0,0), (-1,-1), 2),
    ]))
    story.append(meta_table)
    story.append(Spacer(1, 8))

    # Main table
    table_rows = build_grouped_pdf_rows_compact(df_quote)
    table_rows[0] = [
        Paragraph("<b>Service</b>", head_center),
        Paragraph("<b>Details</b>", head_center),
        Paragraph("<b>Annual Fees</b><br/><b>(Rs.)</b>", head_center),
    ]
    col_widths = [60*mm, 80*mm, 30*mm]
    brand_blue = colors.HexColor(BRAND_BLUE_HEX)
    table = Table(table_rows, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("FONTSIZE", (0,0), (-1,0), 10),
        ("BACKGROUND", (0,0), (-1,0), brand_blue),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("VALIGN", (0,0), (-1,0), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,0), 6),
        ("BOTTOMPADDING", (0,0), (-1,0), 6),
        ("BOX", (0,0), (-1,0), 0.9, brand_blue),
        ("INNERGRID", (0,0), (-1,0), 0.9, brand_blue),
        ("FONTSIZE", (0,1), (-1,-1), 10),
        ("TOPPADDING", (0,1), (-1,-1), 4),
        ("BOTTOMPADDING", (0,1), (-1,-1), 4),
        ("ALIGN", (2,1), (2,-1), "RIGHT"),
        ("INNERGRID", (0,1), (-1,-1), 0.3, colors.HexColor("#d9d9d9")),
        ("BOX", (0,0), (-1,-1), 1.0, colors.black),
    ]))
    story.append(table)
    story.append(Spacer(1, 8))

    # Totals
    tot_lines = [
        ["", "Subtotal", money_inr(subtotal)],
        ["", f"Discount ({discount_pct:.0f}%)", f"- {money_inr(discount_amt)}"] if discount_amt > 0 else ["", "Discount (0%)", money_inr(0)],
        ["", "Taxable Amount", money_inr(subtotal - discount_amt)],
        ["", "GST (18%)", money_inr(gst_amt)],
        ["", "Grand Total", money_inr(grand)],
    ]
    t2 = Table([["","",""], *tot_lines], colWidths=col_widths)
    t2.setStyle(TableStyle([
        ("FONTSIZE", (0,0), (-1,-1), 10),
        ("ALIGN", (2,0), (2,-1), "RIGHT"),
        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
        ("BACKGROUND", (0,-1), (-1,-1), colors.HexColor("#fafafa")),
        ("LINEABOVE", (0,1), (-1,1), 0.5, colors.grey),
        ("TOPPADDING", (0,1), (-1,-1), 4),
        ("BOTTOMPADDING", (0,1), (-1,-1), 4),
        ("BOX", (0,0), (-1,-1), 1.0, colors.black),
    ]))
    story.append(t2)
    story.append(Spacer(1, 8))

    # Notes
    notes = (
        "<b>Note:</b><br/>"
        "1. The fees are exclusive of taxes and out-of-pocket expenses.<br/>"
        "2. GST 18% extra.<br/>"
        "3. Our scope is limited to the services listed above.<br/>"
        "4. The above quotation is valid for a period of 30 days."
    )
    story.append(Paragraph(notes, styles["Normal"]))
    story.append(Spacer(1, 10))

    # Event-based page
    if not df_event.empty:
        story.append(PageBreak())
        story.append(Paragraph("<b>Event-based charges (as applicable, not included in annual fees)</b>", styles["Normal"]))
        story.append(Spacer(1, 4))
        ev_rows = build_event_pdf_rows(df_event)
        ev_rows[0] = [
            Paragraph("<b>Details</b>", head_center),
            Paragraph("<b>Fees</b><br/><b>(Rs.)</b>", head_center),
        ]
        ev = Table(ev_rows, colWidths=[140*mm, 30*mm], repeatRows=1)
        ev.setStyle(TableStyle([
            ("FONTSIZE", (0,0), (-1,0), 10),
            ("BACKGROUND", (0,0), (-1,0), brand_blue),
            ("TEXTCOLOR", (0,0), (-1,0), colors.white),
            ("VALIGN", (0,0), (-1,0), "MIDDLE"),
            ("TOPPADDING", (0,0), (-1,0), 6),
            ("BOTTOMPADDING", (0,0), (-1,0), 6),
            ("BOX", (0,0), (-1,0), 0.9, brand_blue),
            ("INNERGRID", (0,0), (-1,0), 0.9, brand_blue),
            ("FONTSIZE", (0,1), (-1,-1), 10),
            ("TOPPADDING", (0,1), (-1,-1), 4),
            ("BOTTOMPADDING", (0,1), (-1,-1), 4),
            ("ALIGN", (1,1), (1,-1), "RIGHT"),
            ("INNERGRID", (0,1), (-1,-1), 0.3, colors.HexColor("#d9d9d9")),
            ("BOX", (0,0), (-1,-1), 1.0, colors.black),
        ]))
        story.append(ev)
        story.append(Spacer(1, 10))

    doc.build(story)
    return buf.getvalue()

# --------- Excel export ---------
def export_proposal_excel(df_main, df_event, client_name, client_type, quote_no,
                          subtotal, discount_pct, discount_amt, gst_amt, grand):
    import io
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    wb = Workbook()
    thin = Side(style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    head_fill = PatternFill("solid", fgColor=BRAND_BLUE_HEX.replace("#",""))
    head_font = Font(color="FFFFFF", bold=True)
    right = Alignment(horizontal="right")
    center = Alignment(horizontal="center")

    ws = wb.active; ws.title = "Annual Fees"
    headers = ["Service", "Details", "Annual Fees (Rs.)"]
    ws.append(headers)
    for c in range(1, 3+1):
        cell = ws.cell(row=1, column=c)
        cell.fill, cell.font, cell.alignment, cell.border = head_fill, head_font, center, border_all

    for svc, grp in df_main.groupby("Service", sort=True):
        g = grp.sort_values(["Details"]).copy()
        first = True
        for _, r in g.iterrows():
            amt = int(float(r.get("Annual Fees (Rs.)", 0) or 0))
            ws.append([svc if first else "", r.get("Details", ""), amt])
            first = False

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=3):
        row[2].number_format = '#,##,##0'
        for cell in row: cell.border = border_all

    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 52
    ws.column_dimensions["C"].width = 18

    start = ws.max_row + 2
    ws.cell(row=start, column=2, value="Totals").font = Font(bold=True)
    items = [
        ("Subtotal", subtotal),
        (f"Discount ({int(discount_pct)}%)", -discount_amt if discount_amt else 0),
        ("Taxable Amount", subtotal - discount_amt),
        ("GST (18%)", gst_amt),
        ("Grand Total", grand),
    ]
    for i, (lbl, amt) in enumerate(items, start=start+1):
        ws.cell(row=i, column=2, value=lbl)
        c = ws.cell(row=i, column=3, value=int(round(amt)))
        c.number_format = '#,##,##0'; c.alignment = right

    if not df_event.empty:
        ws2 = wb.create_sheet("Event-based charges")
        ws2.append(["Details", "Fees (Rs.)"])
        h1, h2 = ws2["A1"], ws2["B1"]
        for h in (h1, h2):
            h.fill = head_fill; h.font = head_font; h.alignment = center; h.border = border_all
        for _, r in df_event.iterrows():
            detail = (str(r.get("Details","")).strip() or str(r.get("Service","")).strip())
            amt = int(float(r.get("Annual Fees (Rs.)", 0) or 0))
            ws2.append([detail, amt])
        for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row, min_col=1, max_col=2):
            row[1].number_format = '#,##,##0'
            for cell in row: cell.border = border_all
        ws2.column_dimensions["A"].width, ws2.column_dimensions["B"].width = 60, 18

    ws3 = wb.create_sheet("Cover")
    meta = [
        ["Client Name", client_name],
        ["Client Entity Type", client_type],
        ["Quotation No.", quote_no],
        ["Date", datetime.now().strftime("%d-%b-%Y")],
    ]
    for r in meta: ws3.append(r)
    ws3.column_dimensions["A"].width = 22
    ws3.column_dimensions["B"].width = 50

    bio = io.BytesIO(); wb.save(bio)
    return bio.getvalue()

def build_status(df_app, df_fees):
    active = df_app[df_app["Applicable"] == True].copy()
    counts = (
        active.groupby("ClientType")
        .size()
        .reindex([normalize_str(x) for x in CLIENT_TYPES], fill_value=0)
        .reset_index(name="Applicable services")
    )
    merged = active.merge(df_fees, on=["Service", "SubService", "ClientType"], how="left")
    missing_mask = merged["FeeINR"].isna() | (pd.to_numeric(merged["FeeINR"], errors="coerce").fillna(0.0) <= 0)
    miss = (
        merged[missing_mask]
        .groupby("ClientType")
        .size()
        .reindex([normalize_str(x) for x in CLIENT_TYPES], fill_value=0)
        .reset_index(name="Missing/Zero fees")
    )
    status = counts.merge(miss, on="ClientType")
    status["ClientType"] = status["ClientType"].str.title()
    return status

# --------- UI ---------
st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="centered")

try:
    df_app, df_fees, source = load_matrices()
except Exception as e:
    st.error(f"Error loading matrices: {e}")
    st.stop()

with st.sidebar:
    st.markdown(
        "**What this tool does:**\n\n"
        "1. Create annual fees quotations based on client type\n"
        "2. Fees derived from fees master already provided\n"
        "3. Fees amount is still editable\n"
        "4. Export to PDF/Excel"
    )
    st.divider()
    st.subheader("Options")
    st.session_state["discount_pct"] = st.number_input("1) Discount %", 0, 100, int(st.session_state["discount_pct"]), 1)
    st.session_state["letterhead"] = st.checkbox("2) Letterhead mode (watermark logo)", value=st.session_state["letterhead"])
    sig_up = st.file_uploader("3) Signature / Stamp image (optional)", type=["png", "jpg", "jpeg"])
    if sig_up is not None:
        st.session_state["sig_bytes"] = sig_up.read()
    st.divider()
    with st.expander("Data status", expanded=False):
        service_defs = len(df_app[["Service", "SubService"]].drop_duplicates())
        st.write(f"Source: **{source}**")
        st.write(f"Service definitions: **{service_defs}**")
        status_df = build_status(df_app, df_fees)
        all_ok = (status_df["Missing/Zero fees"] == 0).all()
        if all_ok:
            st.success("All applicable services have fees.")
            st.dataframe(status_df.drop(columns=["Missing/Zero fees"]), use_container_width=True)
        else:
            st.warning("Some fees are missing/zero. Review below.")
            st.dataframe(status_df, use_container_width=True)

st.markdown(f"""
<style>
.stApp {{ background: #ffffff !important; }}
html, body {{ color: #111 !important; }}
.stButton > button, .stDownloadButton > button {{
  background: {BRAND_BLUE_HEX} !important; color: #fff !important; border: 0; border-radius: 6px;
}}
h1 > a, h2 > a, h3 > a, h4 > a {{ display: none !important; }}
.label-lg {{ font-size: 1.05rem; font-weight: 700; margin: 6px 0 2px 0; }}
</style>
""", unsafe_allow_html=True)

st.title("Quotation Generator")
st.subheader("V. Purohit & Associates")

# --------- FORM ---------
with st.form("quote_form", clear_on_submit=False):
    st.markdown('<div class="label-lg">➤ Client Name*</div>', unsafe_allow_html=True)
    client_name = st.text_input("", st.session_state.get("client_name", ""), placeholder="Enter client name")

    st.markdown('<div class="label-lg">➤ Client Type*</div>', unsafe_allow_html=True)
    client_type = st.selectbox("", CLIENT_TYPES, index=0)

    st.markdown('<div class="label-lg">➤ Client Contact (optional)</div>', unsafe_allow_html=True)
    addr = st.text_area("", st.session_state.get("client_addr", ""), placeholder="Street, Area\nCity, State, PIN")
    colX, colY = st.columns(2)
    with colX:
        email = st.text_input("", st.session_state.get("client_email", ""), placeholder="Email")
    with colY:
        phone = st.text_input("", st.session_state.get("client_phone", ""), placeholder="Phone")

    st.markdown('<div class="label-lg">➤ Proposed Start (optional)</div>', unsafe_allow_html=True)
    proposal_start = st.text_input("", st.session_state.get("proposal_start", ""), placeholder="e.g., Aug 2025 / FY 2025-26")

    ct_norm = normalize_str(client_type)
    app_ct = df_app[(df_app["ClientType"] == ct_norm) & (df_app["Applicable"] == True)]

    st.markdown('<div class="label-lg">➤ Accounting – choose one plan</div>', unsafe_allow_html=True)
    selected_accounting = st.radio("", ACCOUNTING_PLANS, index=3, horizontal=True)

    st.markdown('<div class="label-lg">➤ Profession Tax Returns – choose one type</div>', unsafe_allow_html=True)
    pt_options = app_ct.loc[app_ct["Service"] == normalize_str(PT_SERVICE), "SubService"].dropna().unique().tolist()
    pt_options_tc = sorted([title_with_acronyms(s) for s in pt_options if s])
    selected_pt_tc = st.radio("", pt_options_tc if pt_options_tc else ["(Not applicable)"], index=0, horizontal=True)
    if selected_pt_tc == "(Not applicable)":
        selected_pt_tc = None

    submit = st.form_submit_button("Generate Table")

# --------- Build dataframes on submit ---------
if submit:
    if not client_name.strip():
        st.error("Please enter Client Name.")
    else:
        st.session_state["quote_no"] = datetime.now().strftime("QTN-%Y%m%d-%H%M%S")
        st.session_state["proposal_start"] = proposal_start
        main_df, event_df, _ = build_quotes(
            client_name, client_type, df_app, df_fees,
            selected_accounting=selected_accounting,
            selected_pt_sub=selected_pt_tc,
        )
        if main_df.empty and event_df.empty:
            st.warning("No applicable services found for the selected Client Type.")
        else:
            df_main = main_df.copy()
            df_main["Include"] = True
            df_main["MoveToEvent"] = False
            st.session_state["quote_df"] = df_main
            st.session_state["event_df"] = event_df.copy()
            st.session_state["client_name"] = client_name
            st.session_state["client_type"] = client_type
            st.session_state["client_addr"] = addr
            st.session_state["client_email"] = email
            st.session_state["client_phone"] = phone
            st.session_state["editor_active"] = True

# --------- Editors / Totals / Exports ---------
if st.session_state["editor_active"] and (
    not st.session_state["quote_df"].empty or not st.session_state["event_df"].empty
):
    st.success("Edit annual fees below. Event-based charges are listed separately and not included in totals.")

    # MAIN editor – Include, MoveToEvent, fees
    if not st.session_state["quote_df"].empty:
        with st.form("edit_main"):
            edited = st.data_editor(
                st.session_state["quote_df"],
                use_container_width=True,
                disabled=["Service", "Details"],
                column_order=["Include", "MoveToEvent", "Service", "Details", "Annual Fees (Rs.)"],
                column_config={
                    "Include": st.column_config.CheckboxColumn(help="Uncheck to remove this row from the proposal/PDF."),
                    "MoveToEvent": st.column_config.CheckboxColumn(help="Tick and submit to shift to Event-based table."),
                    "Annual Fees (Rs.)": st.column_config.NumberColumn(
                        "Annual Fees (Rs.)", min_value=0, step=100, format="%.0f",
                        help="Edit the fee; totals & PDF will use this value."
                    ),
                },
                num_rows="fixed", key="quote_editor", hide_index=True, height=420,
            )
            col_m1, col_m2 = st.columns([1,1])
            apply_edits = col_m1.form_submit_button("Apply edits")
            apply_and_move = col_m2.form_submit_button("Apply edits & move selected")
        if apply_edits or apply_and_move:
            st.session_state["quote_df"] = edited
            if apply_and_move:
                move_rows = edited[edited["MoveToEvent"] == True].copy()
                if not move_rows.empty:
                    addon = move_rows[["Service", "Details", "Annual Fees (Rs.)"]].copy()
                    addon["Details"] = addon["Details"].apply(lambda x: x.strip() if isinstance(x, str) else "")
                    addon.loc[addon["Details"] == "", "Details"] = addon["Service"]
                    # add to event
                    ev_now = st.session_state["event_df"].copy()
                    ev_now = pd.concat([ev_now, addon], ignore_index=True)
                    st.session_state["event_df"] = ev_now
                    kept = edited.loc[edited["MoveToEvent"] != True].copy()
                    kept["MoveToEvent"] = False
                    st.session_state["quote_df"] = kept
                    st.success(f"Moved {len(addon)} row(s) to Event-based.")

        qdf = st.session_state["quote_df"]
        filtered = qdf[qdf["Include"] == True].copy() if "Include" in qdf.columns else qdf.copy()
        for col in ["Include", "MoveToEvent"]:
            if col in filtered.columns:
                filtered.drop(columns=[col], inplace=True)
        filtered["Annual Fees (Rs.)"] = pd.to_numeric(filtered["Annual Fees (Rs.)"], errors="coerce").fillna(0.0)
    else:
        filtered = pd.DataFrame(columns=["Service", "Details", "Annual Fees (Rs.)"])

    # EVENT editor – fees editable (no move-back here)
    event_df = st.session_state["event_df"].copy()
    if not event_df.empty:
        st.subheader("Event-based charges (as applicable)")
        st.caption("These are not included in the annual fees totals.")
        with st.form("event_form"):
            event_edited = st.data_editor(
                event_df,
                use_container_width=True,
                disabled=["Service", "Details"],
                column_order=["Service", "Details", "Annual Fees (Rs.)"],
                column_config={
                    "Annual Fees (Rs.)": st.column_config.NumberColumn(
                        "Fees (Rs.)", min_value=0, step=100, format="%.0f",
                        help="Edit the fee; shown separately and not included in totals."
                    ),
                },
                num_rows="fixed", key="event_editor", hide_index=True, height=320,
            )
            ev_apply = st.form_submit_button("Apply event edits")
        if ev_apply:
            st.session_state["event_df"] = event_edited

    # Totals (UI)
    subtotal, discount_amt, taxable, gst_amt, grand = compute_totals(filtered, st.session_state["discount_pct"])
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        st.write(f"**Subtotal (Rs.):** {money_inr(subtotal)}")
        st.write(f"**Discount ({st.session_state['discount_pct']}%) (Rs.):** {money_inr(discount_amt)}")
    with c2:
        st.write(f"**Taxable Amount (Rs.):** {money_inr(taxable)}")
        st.write(f"**GST (18%) (Rs.):** {money_inr(gst_amt)}")
    with c3:
        st.write(f"**Grand Total (Rs.):** {money_inr(grand)}")
        if st.button("Start Over"):
            st.session_state["editor_active"] = False
            st.session_state["quote_df"] = pd.DataFrame()
            st.session_state["event_df"] = pd.DataFrame()
            st.rerun()

    # Excel
    if not filtered.empty:
        try:
            excel_bytes = export_proposal_excel(
                filtered, st.session_state["event_df"],
                st.session_state["client_name"], st.session_state["client_type"],
                st.session_state.get("quote_no",""),
                subtotal, float(st.session_state["discount_pct"]), discount_amt, gst_amt, grand
            )
            st.download_button(
                "⬇️ Download Excel (Full Proposal)",
                data=excel_bytes,
                file_name="Annual_Fees_Proposal.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_xlsx_full",
            )
        except Exception:
            st.warning("Excel export requires 'openpyxl' on the server.")

    # PDF
    pdf_bytes = make_pdf(
        st.session_state["client_name"], st.session_state["client_type"],
        datetime.now().strftime("QTN-%Y%m%d-%H%M%S") if not st.session_state.get("quote_no") else st.session_state["quote_no"],
        filtered, st.session_state["event_df"],
        subtotal, float(st.session_state["discount_pct"]), discount_amt, gst_amt, grand,
        letterhead=st.session_state["letterhead"],
        addr=st.session_state.get("client_addr", ""), email=st.session_state.get("client_email", ""),
        phone=st.session_state.get("client_phone", ""), proposal_start=st.session_state.get("proposal_start", ""),
    )
    st.download_button(
        "⬇️ Download PDF",
        data=pdf_bytes,
        file_name=f"Annual_Fees_Proposal_{st.session_state['client_name'].replace(' ', '_')}.pdf",
        mime="application/pdf", key="dl_pdf",
    )
