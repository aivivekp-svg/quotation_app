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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

APP_TITLE = "Quotation Generator – V. Purohit & Associates"

CLIENT_TYPES = [
    "PRIVATE LIMITED", "PROPRIETORSHIP", "INDIVIDUAL", "LLP", "HUF",
    "SOCIETY", "PARTNERSHIP FIRM", "FOREIGN ENTITY", "AOP/ BOI", "TRUST",
]

ACCOUNTING_PLANS = ["Monthly Accounting", "Quarterly Accounting", "Half Yearly Accounting", "Annual Accounting"]

EVENT_SERVICE = "EVENT BASED FILING"
PT_SERVICE = "PROFESSION TAX RETURNS"

FIRM_FOOTER = (
    "Office No. 5, Ground Floor, Adeshwar Arcade Commercial Premises CSL, "
    "Andheri - Kurla Road, Opp. Sangam Cinema, Andheri East, Mumbai - 400093.  "
    "Email: info@vpurohit.com, Contact: +91-8369508539"
)

ACRONYMS = ["GST", "GSTR", "PTEC", "PTRC", "ADT", "ROC", "TDS", "AOC", "MGT", "26QB", "26QC"]

GST_RATE_FIXED = 18  # always 18%

# ---------- Session defaults ----------
if "editor_active" not in st.session_state: st.session_state["editor_active"] = False
if "quote_df" not in st.session_state: st.session_state["quote_df"] = pd.DataFrame()
if "client_name" not in st.session_state: st.session_state["client_name"] = ""
if "client_type" not in st.session_state: st.session_state["client_type"] = ""
if "quote_no" not in st.session_state: st.session_state["quote_no"] = ""
if "discount_pct" not in st.session_state: st.session_state["discount_pct"] = 0
if "letterhead" not in st.session_state: st.session_state["letterhead"] = False
if "client_addr" not in st.session_state: st.session_state["client_addr"] = ""
if "client_email" not in st.session_state: st.session_state["client_email"] = ""
if "client_phone" not in st.session_state: st.session_state["client_phone"] = ""
if "sig_bytes" not in st.session_state: st.session_state["sig_bytes"] = None  # signature image bytes

# ---------- Helpers ----------
def normalize_str(x):
    return (x or "").strip().upper()

def title_with_acronyms(text: str) -> str:
    """Title-case, fix 'of' to lower-case, and force known acronyms to uppercase."""
    if text is None: return ""
    t = " ".join(str(text).split()).title()
    t = re.sub(r"\bOf\b", "of", t, flags=re.IGNORECASE)
    for token in ACRONYMS:
        t = re.sub(rf"\b{re.escape(token)}\b", token, t, flags=re.IGNORECASE)
    return t

def service_display_override(raw_upper: str, pretty: str) -> str:
    if raw_upper == "FILING OF GSTR RETURNS":
        return "Filing of GST Returns"
    return pretty

def money(n: float) -> str:
    return f"{n:,.0f}"

def load_matrices():
    xl = pd.ExcelFile("matrices.xlsx")  # always from repo
    df_app = xl.parse("Applicability").fillna("")
    df_fees = xl.parse("Fees").fillna("")
    for df in (df_app, df_fees):
        df["Service"] = df["Service"].map(normalize_str)
        df["SubService"] = df["SubService"].map(lambda v: normalize_str(v) if pd.notna(v) else "")
        df["ClientType"] = df["ClientType"].map(normalize_str)
    df_app["Applicable"] = df_app["Applicable"].astype(str).str.upper().isin(["TRUE","1","YES"])
    df_fees["FeeINR"] = pd.to_numeric(df_fees["FeeINR"], errors="coerce").fillna(0.0).astype(float)
    return df_app, df_fees, "matrices.xlsx"

def build_quote(client_name, client_type, df_app, df_fees,
                selected_accounting=None, selected_event_subs=None, selected_pt_sub=None):
    ct = normalize_str(client_type)
    applicable = (
        df_app.query("ClientType == @ct and Applicable == True")
              .loc[:, ["Service","SubService","ClientType"]]
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
    # Event Based Filing → multiselect
    is_event = applicable["Service"].eq(normalize_str(EVENT_SERVICE))
    if selected_event_subs is not None:
        if len(selected_event_subs) == 0:
            applicable = applicable.loc[~is_event]
        else:
            sel_set = {normalize_str(s) for s in selected_event_subs}
            applicable = pd.concat(
                [applicable.loc[~is_event], applicable.loc[is_event & (applicable["SubService"].isin(sel_set))]],
                ignore_index=True,
            )
    # Profession Tax Returns → choose one
    is_pt = applicable["Service"].eq(normalize_str(PT_SERVICE))
    if selected_pt_sub is not None:
        sel_pt = normalize_str(selected_pt_sub)
        applicable = pd.concat(
            [applicable.loc[~is_pt], applicable.loc[is_pt & (applicable["SubService"] == sel_pt)]],
            ignore_index=True,
        )
    # Merge with Fees
    quoted = applicable.merge(
        df_fees, on=["Service","SubService","ClientType"], how="left", validate="1:1"
    )
    quoted["FeeINR"] = pd.to_numeric(quoted["FeeINR"], errors="coerce").fillna(0.0)
    # Labels with acronyms + overrides
    service_key_upper = quoted["Service"].copy()
    svc_pretty = service_key_upper.map(title_with_acronyms)
    quoted["Service"] = [
        service_display_override(raw, pretty)
        for raw, pretty in zip(service_key_upper.tolist(), svc_pretty.tolist())
    ]
    quoted["SubService"] = quoted["SubService"].map(title_with_acronyms)
    quoted.sort_values(["Service","SubService"], inplace=True)
    out = (quoted.drop(columns=["ClientType"], errors="ignore")
                 .rename(columns={"SubService":"Details","FeeINR":"Annual Fees (Rs.)"})
                 .loc[:, ["Service","Details","Annual Fees (Rs.)"]])
    total = float(out["Annual Fees (Rs.)"].sum()) if not out.empty else 0.0
    return out, total

def compute_totals(df_selected: pd.DataFrame, discount_pct: float):
    fees_series = pd.to_numeric(df_selected["Annual Fees (Rs.)"], errors="coerce").fillna(0.0)
    subtotal = float(fees_series.sum())
    discount_amt = round(subtotal * (discount_pct or 0) / 100.0, 2)
    taxable = max(subtotal - discount_amt, 0.0)
    gst_amt = round(taxable * GST_RATE_FIXED / 100.0, 2)
    grand = round(taxable + gst_amt, 2)
    return subtotal, discount_amt, taxable, gst_amt, grand

# --- PDF table rows: grouped with a single service-name row (no HTML tags) ---
def build_grouped_pdf_rows(df: pd.DataFrame):
    rows = [["Service", "Details", "Annual Fees<br/>(Rs.)"]]
    styles = []
    r = 1
    for svc, grp in df.groupby("Service", sort=True):
        rows.append([svc, "", ""])
        styles.extend([
            ("BACKGROUND", (0, r), (-1, r), colors.HexColor("#f7f7f7")),
            ("FONTNAME", (0, r), (-1, r), "Helvetica-Bold"),
            ("ALIGN", (0, r), (-1, r), "LEFT"),
        ])
        r += 1
        for _, row in grp.iterrows():
            amt = pd.to_numeric(row["Annual Fees (Rs.)"], errors="coerce")
            amt = 0.0 if pd.isna(amt) else float(amt)
            rows.append(["", row["Details"], money(amt)])
            r += 1
    return rows, styles

def make_pdf(client_name: str, client_type: str, quote_no: str, df_quote: pd.DataFrame,
             subtotal: float, discount_pct: float, discount_amt: float, gst_amt: float, grand: float,
             letterhead: bool = False, addr: Optional[str] = None, email: Optional[str] = None,
             phone: Optional[str] = None, signature_bytes: Optional[bytes] = None):
    import os
    from reportlab.platypus import Image
    from reportlab.lib.utils import ImageReader

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=16*mm, bottomMargin=16*mm
    )
    styles = getSampleStyleSheet()
    head_center = ParagraphStyle("HeadCenter", parent=styles["Normal"], alignment=TA_CENTER, fontName="Helvetica-Bold")
    story = []

    # Logo (small; watermark handled in decorator if letterhead)
    def find_logo_path():
        for name in ("logo.png","logo.jpg","logo.jpeg"):
            if os.path.exists(name): return name
        return None
    logo_path = find_logo_path()
    if logo_path and not letterhead:
        ir = ImageReader(logo_path)
        ow, oh = ir.getSize(); max_w, max_h = 26*mm, 26*mm
        r = min(max_w/ow, max_h/oh)
        story.append(Image(logo_path, width=ow*r, height=oh*r))
        story.append(Spacer(1, 4))

    # Title
    story.append(Paragraph("<b>V. Purohit & Associates</b>", styles["Title"]))
    story.append(Paragraph("<b>Annual Fees Proposal</b>", styles["h2"]))
    story.append(Spacer(1, 6))

    # Meta info (requested order)
    meta_lines = [
        f"<b>Client Name:</b> {client_name}",
        f"<b>Client Type:</b> {client_type}",
        f"<b>Quotation No.:</b> {quote_no}",
        f"<b>Date:</b> {datetime.now().strftime('%d-%b-%Y')}",
    ]
    if addr and addr.strip():
        addr_html = "<br/>".join([ln.strip() for ln in addr.splitlines() if ln.strip()])
        meta_lines.append(f"<b>Address:</b> {addr_html}")
    if email and email.strip():
        meta_lines.append(f"<b>Email:</b> {email.strip()}")
    if phone and phone.strip():
        meta_lines.append(f"<b>Phone:</b> {phone.strip()}")
    story.append(Paragraph("<br/>".join(meta_lines), styles["Normal"]))
    story.append(Spacer(1, 8))

    # Table (reverted widths: 60mm / 80mm / 30mm)
    table_rows, extra_styles = build_grouped_pdf_rows(df_quote)
    table_rows[0] = [
        Paragraph("<b>Service</b>", head_center),
        Paragraph("<b>Details</b>", head_center),
        Paragraph("<b>Annual Fees</b><br/><b>(Rs.)</b>", head_center),
    ]
    col_widths = [60*mm, 80*mm, 30*mm]
    table = Table(table_rows, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle([
        # Header row with borders between headings + box
        ("FONTSIZE", (0,0), (-1,0), 10),
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f0f0f0")),
        ("VALIGN", (0,0), (-1,0), "MIDDLE"),
        ("TOPPADDING", (0,0), (-1,0), 6),
        ("BOTTOMPADDING", (0,0), (-1,0), 6),
        ("BOX", (0,0), (-1,0), 0.9, colors.grey),
        ("INNERGRID", (0,0), (-1,0), 0.9, colors.grey),
        # Body
        ("FONTSIZE", (0,1), (-1,-1), 10),
        ("TOPPADDING", (0,1), (-1,-1), 4),
        ("BOTTOMPADDING", (0,1), (-1,-1), 4),
        ("ALIGN", (2,1), (2,-1), "RIGHT"),
        ("INNERGRID", (0,1), (-1,-1), 0.3, colors.HexColor("#d9d9d9")),
        # Outer border around whole table
        ("BOX", (0,0), (-1,-1), 1.0, colors.black),
    ] + extra_styles))
    story.append(table)
    story.append(Spacer(1, 8))

    # Totals block (same widths; boxed)
    tot_lines = [
        ["", "Subtotal", money(subtotal)],
        ["", f"Discount ({discount_pct:.0f}%)", f"- {money(discount_amt)}"] if discount_amt > 0 else ["", "Discount (0%)", money(0)],
        ["", "Taxable Amount", money(subtotal - discount_amt)],
        ["", "GST (18%)", money(gst_amt)],
        ["", "Grand Total", money(grand)],
    ]
    t2 = Table([["","", ""], *tot_lines], colWidths=col_widths)
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
    story.append(Spacer(1, 10))

    # Signature / stamp block (optional)
    if signature_bytes:
        try:
            from reportlab.platypus import Image
            story.append(Paragraph("<b>For V. Purohit & Associates</b>", styles["Normal"]))
            story.append(Spacer(1, 4))
            story.append(Image(io.BytesIO(signature_bytes), width=45*mm, height=20*mm))
            story.append(Paragraph("Authorised Signatory", styles["Normal"]))
            story.append(Spacer(1, 6))
        except Exception:
            pass
    else:
        for alt in ("signature.png", "signature.jpg", "stamp.png", "stamp.jpg"):
            try:
                import os
                if os.path.exists(alt):
                    from reportlab.platypus import Image
                    story.append(Paragraph("<b>For V. Purohit & Associates</b>", styles["Normal"]))
                    story.append(Spacer(1, 4))
                    story.append(Image(alt, width=45*mm, height=20*mm))
                    story.append(Paragraph("Authorised Signatory", styles["Normal"]))
                    story.append(Spacer(1, 6))
                    break
            except Exception:
                continue

    # Footer / page numbers / optional watermark
    def _decorate(canv, doc_):
        if letterhead:
            try:
                import os
                from reportlab.lib.utils import ImageReader
                for n in ("logo.png","logo.jpg","logo.jpeg"):
                    if os.path.exists(n):
                        canv.saveState()
                        if hasattr(canv, "setFillAlpha"): canv.setFillAlpha(0.07)
                        ir = ImageReader(n)
                        ow, oh = ir.getSize(); w = A4[0] - 60*mm; r = w / ow; h = oh * r
                        x = 30*mm; y = (A4[1] - h)/2
                        canv.drawImage(n, x, y, width=w, height=h, preserveAspectRatio=True, mask="auto")
                        canv.restoreState()
                        break
            except Exception:
                pass

        canv.saveState()
        canv.setFont("Helvetica", 8)
        # Line above footer
        y_line = 20*mm
        canv.setLineWidth(0.7)
        canv.line(18*mm, y_line, A4[0] - 18*mm, y_line)
        # Footer text
        line1 = "Office No. 5, Ground Floor, Adeshwar Arcade Commercial Premises CSL, Andheri - Kurla Road,"
        line2 = "Opp. Sangam Cinema, Andheri East, Mumbai - 400093.  Email: info@vpurohit.com, Contact: +91-8369508539"
        canv.drawCentredString(A4[0] / 2, 12*mm + 4, line1)
        canv.drawCentredString(A4[0] / 2, 12*mm - 6, line2)
        canv.drawRightString(A4[0] - 18*mm, 12*mm - 18, f"Page {canv.getPageNumber()}")
        canv.restoreState()

    doc.build(story, onFirstPage=_decorate, onLaterPages=_decorate)
    return buf.getvalue()

def build_status(df_app, df_fees):
    active = df_app[df_app["Applicable"] == True].copy()
    counts = (active.groupby("ClientType").size()
              .reindex([normalize_str(x) for x in CLIENT_TYPES], fill_value=0)
              .reset_index(name="Applicable services"))
    merged = active.merge(df_fees, on=["Service","SubService","ClientType"], how="left")
    missing_mask = merged["FeeINR"].isna() | (pd.to_numeric(merged["FeeINR"], errors="coerce").fillna(0.0) <= 0)
    miss = (merged[missing_mask].groupby("ClientType").size()
            .reindex([normalize_str(x) for x in CLIENT_TYPES], fill_value=0)
            .reset_index(name="Missing/Zero fees"))
    status = counts.merge(miss, on="ClientType")
    status["ClientType"] = status["ClientType"].str.title()
    return status

# ---------- UI ----------
st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="centered")

# Two-line header in UI
st.title("Quotation Generator")
st.subheader("V. Purohit & Associates")

with st.sidebar:
    st.markdown(
        "**What this tool does:**\n\n"
        "1. Create annual fees quotations based on client type\n"
        "2. Fees derived from fees master already provided\n"
        "3. Fees amount is still editable\n"
        "4. Export to PDF/Excel/CSV"
    )
    st.divider()


    # Numbered options
    st.subheader("Options")
    st.session_state["discount_pct"] = st.number_input("1) Discount %", 0, 100, int(st.session_state["discount_pct"]), 1)
    st.session_state["letterhead"] = st.checkbox("2) Letterhead mode (watermark logo)", value=st.session_state["letterhead"])
    sig_up = st.file_uploader("3) Signature / Stamp image (optional)", type=["png","jpg","jpeg"])
    if sig_up is not None:
        st.session_state["sig_bytes"] = sig_up.read()

    # Load matrices (no upload option)
    try:
        df_app, df_fees, source = load_matrices()
        with st.expander("Data status", expanded=False):
            service_defs = len(df_app[["Service","SubService"]].drop_duplicates())
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
    except Exception as e:
        st.error(f"Error loading matrices: {e}")
        st.stop()

# ----- Form to generate initial table -----
with st.form("quote_form", clear_on_submit=False):
    client_name = st.text_input("Client Name*", st.session_state.get("client_name",""))
    client_type = st.selectbox("Client Type*", CLIENT_TYPES, index=0)

    # Optional client contact fields (used in PDF if provided)
    st.markdown("**Client Contact (optional, for PDF header)**")
    addr = st.text_area("Address", st.session_state.get("client_addr",""), placeholder="Street, Area\nCity, State, PIN")
    email = st.text_input("Email", st.session_state.get("client_email",""))
    phone = st.text_input("Phone", st.session_state.get("client_phone",""))

    ct_norm = normalize_str(client_type)
    app_ct = df_app[(df_app["ClientType"] == ct_norm) & (df_app["Applicable"] == True)]

    selected_accounting = st.radio("Accounting – choose one plan", ACCOUNTING_PLANS, index=3, horizontal=True)

    event_options = app_ct.loc[app_ct["Service"] == normalize_str(EVENT_SERVICE), "SubService"].dropna().unique().tolist()
    event_options_tc = sorted([title_with_acronyms(s) for s in event_options if s])
    selected_event_tc = st.multiselect(
        f"{title_with_acronyms(EVENT_SERVICE)} – select sub-services (choose any)",
        event_options_tc, default=[], help="Only selected items will be included in the proposal."
    )

    pt_options = app_ct.loc[app_ct["Service"] == normalize_str(PT_SERVICE), "SubService"].dropna().unique().tolist()
    pt_options_tc = sorted([title_with_acronyms(s) for s in pt_options if s])
    selected_pt_tc = st.radio(
        f"{title_with_acronyms(PT_SERVICE)} – choose one",
        pt_options_tc if pt_options_tc else ["(Not applicable)"], index=0, horizontal=True
    )
    if selected_pt_tc == "(Not applicable)": selected_pt_tc = None

    submit = st.form_submit_button("Generate Table")

# ----- Persisted editor: simple table (no grouping), fees editable, include toggle -----
if submit:
    if not client_name.strip():
        st.error("Please enter Client Name.")
    else:
        st.session_state["quote_no"] = datetime.now().strftime("QTN-%Y%m%d-%H%M%S")
        df_quote, _ = build_quote(
            client_name, client_type, df_app, df_fees,
            selected_accounting=selected_accounting,
            selected_event_subs=selected_event_tc,
            selected_pt_sub=selected_pt_tc,
        )
        if df_quote.empty:
            st.warning("No applicable services found for the selected Client Type.")
        else:
            df_quote = df_quote.copy()
            df_quote["Include"] = True
            st.session_state["quote_df"] = df_quote
            st.session_state["client_name"] = client_name
            st.session_state["client_type"] = client_type
            st.session_state["client_addr"] = addr
            st.session_state["client_email"] = email
            st.session_state["client_phone"] = phone
            st.session_state["editor_active"] = True

if st.session_state["editor_active"] and not st.session_state["quote_df"].empty:
    st.success("Proposal ready. Edit fees or uncheck rows; totals update live.")
    edited = st.data_editor(
        st.session_state["quote_df"],
        use_container_width=True,
        disabled=["Service","Details"],  # Fee is editable
        column_config={
            "Include": st.column_config.CheckboxColumn(help="Uncheck to remove this row from the proposal/PDF."),
            "Annual Fees (Rs.)": st.column_config.NumberColumn(
                "Annual Fees (Rs.)", min_value=0, step=100,
                help="Edit the fee; totals & PDF will use this value.", format="%.0f"
            ),
        },
        num_rows="fixed",
        key="quote_editor",
    )
    st.session_state["quote_df"] = edited

    filtered = edited[edited["Include"] == True].drop(columns=["Include"])
    filtered["Annual Fees (Rs.)"] = pd.to_numeric(filtered["Annual Fees (Rs.)"], errors="coerce").fillna(0.0)

    subtotal, discount_amt, taxable, gst_amt, grand = compute_totals(
        filtered, st.session_state["discount_pct"]
    )

    c1, c2, c3 = st.columns([1,1,1])
    with c1:
        st.write(f"**Subtotal (Rs.):** {money(subtotal)}")
        st.write(f"**Discount ({st.session_state['discount_pct']}%) (Rs.):** {money(discount_amt)}")
    with c2:
        st.write(f"**Taxable Amount (Rs.):** {money(taxable)}")
        st.write(f"**GST (18%) (Rs.):** {money(gst_amt)}")
    with c3:
        st.write(f"**Grand Total (Rs.):** {money(grand)}")
        if st.button("Start Over"):
            st.session_state["editor_active"] = False
            st.session_state["quote_df"] = pd.DataFrame()
            st.rerun()

    if filtered.empty:
        st.info("All rows are excluded. Select at least one row to enable exports/PDF.")
    else:
        # Exports
        csv_bytes = filtered.to_csv(index=False).encode("utf-8")
        xlsx_buf = io.BytesIO()
        try:
            with pd.ExcelWriter(xlsx_buf, engine="openpyxl") as writer:
                filtered.to_excel(writer, index=False, sheet_name="Proposal")
            xlsx_data = xlsx_buf.getvalue()
        except Exception:
            xlsx_data = None

        colA, colB, colC = st.columns([1,1,1])
        with colA:
            st.download_button("⬇️ Download CSV", data=csv_bytes, file_name="annual_fees_proposal.csv", mime="text/csv", key="dl_csv")
        with colB:
            if xlsx_data:
                st.download_button("⬇️ Download Excel", data=xlsx_data, file_name="annual_fees_proposal.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_xlsx")
            else:
                st.caption(":grey[Excel export unavailable (missing engine).]")

        # PDF
        pdf_bytes = make_pdf(
            st.session_state["client_name"], st.session_state["client_type"], st.session_state["quote_no"],
            filtered, subtotal, float(st.session_state["discount_pct"]), discount_amt, gst_amt, grand,
            letterhead=st.session_state["letterhead"],
            addr=st.session_state.get("client_addr",""),
            email=st.session_state.get("client_email",""),
            phone=st.session_state.get("client_phone",""),
            signature_bytes=st.session_state.get("sig_bytes")
        )
        with colC:
            st.download_button(
                "⬇️ Download PDF",
                data=pdf_bytes,
                file_name=f"Annual_Fees_Proposal_{st.session_state['client_name'].replace(' ', '_')}.pdf",
                mime="application/pdf",
                key="dl_pdf"
            )
