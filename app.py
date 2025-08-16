import io
import re
from datetime import datetime

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
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

# Acronyms – also ensures GST is uppercased
ACRONYMS = ["GST", "GSTR", "PTEC", "PTRC", "ADT", "ROC", "TDS", "AOC", "MGT", "26QB", "26QC"]

# ---------- Session defaults ----------
if "editor_active" not in st.session_state: st.session_state["editor_active"] = False
if "rows" not in st.session_state: st.session_state["rows"] = []  # grouped editor rows
if "client_name" not in st.session_state: st.session_state["client_name"] = ""
if "client_type" not in st.session_state: st.session_state["client_type"] = ""
if "quote_no" not in st.session_state: st.session_state["quote_no"] = ""
if "gst_pct" not in st.session_state: st.session_state["gst_pct"] = 18
if "discount_pct" not in st.session_state: st.session_state["discount_pct"] = 0
if "letterhead" not in st.session_state: st.session_state["letterhead"] = False

# ---------- Helpers ----------
def normalize_str(x): return (x or "").strip().upper()

def title_with_acronyms(text: str) -> str:
    """Title-case, fix 'of' to lower-case, and force known acronyms to uppercase."""
    if text is None: return ""
    t = " ".join(str(text).split()).title()
    t = re.sub(r"\bOf\b", "of", t, flags=re.IGNORECASE)  # 'of' → lower
    for token in ACRONYMS:
        t = re.sub(rf"\b{re.escape(token)}\b", token, t, flags=re.IGNORECASE)
    return t

def service_display_override(raw_upper: str, pretty: str) -> str:
    # 'FILING OF GSTR RETURNS' → 'Filing of GST Returns' (service display only)
    if raw_upper == "FILING OF GSTR RETURNS":
        return "Filing of GST Returns"
    return pretty

def money(n: float) -> str: return f"{n:,.0f}"

def load_matrices(uploaded_file):
    if uploaded_file is not None:
        xl = pd.ExcelFile(uploaded_file); source = "Uploaded file"
    else:
        xl = pd.ExcelFile("matrices.xlsx"); source = "matrices.xlsx"

    df_app = xl.parse("Applicability").fillna("")
    df_fees = xl.parse("Fees").fillna("")

    for df in (df_app, df_fees):
        df["Service"] = df["Service"].map(normalize_str)
        df["SubService"] = df["SubService"].map(lambda v: normalize_str(v) if pd.notna(v) else "")
        df["ClientType"] = df["ClientType"].map(normalize_str)

    df_app["Applicable"] = df_app["Applicable"].astype(str).str.upper().isin(["TRUE","1","YES"])
    df_fees["FeeINR"] = pd.to_numeric(df_fees["FeeINR"], errors="coerce").fillna(0.0).astype(float)
    return df_app, df_fees, source

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

    # Presentable labels with acronyms + display overrides
    service_key = quoted["Service"].copy()  # UPPER
    svc_pretty = service_key.map(title_with_acronyms)
    svc_final = [service_display_override(raw, pretty) for raw, pretty in zip(service_key.tolist(), svc_pretty.tolist())]
    quoted["Service"] = svc_final
    quoted["SubService"] = quoted["SubService"].map(title_with_acronyms)

    quoted.sort_values(["Service","SubService"], inplace=True)

    out = (quoted.drop(columns=["ClientType"], errors="ignore")
                 .rename(columns={"SubService":"Details","FeeINR":"Annual Fees (Rs.)"})
                 .loc[:, ["Service","Details","Annual Fees (Rs.)"]])
    total = float(out["Annual Fees (Rs.)"].sum()) if not out.empty else 0.0
    return out, total

def compute_totals(df_selected: pd.DataFrame, discount_pct: float, gst_pct: float):
    fees_series = pd.to_numeric(df_selected["Annual Fees (Rs.)"], errors="coerce").fillna(0.0)
    subtotal = float(fees_series.sum())
    discount_amt = round(subtotal * (discount_pct or 0) / 100.0, 2)
    taxable = max(subtotal - discount_amt, 0.0)
    gst_amt = round(taxable * (gst_pct or 0) / 100.0, 2)
    grand = round(taxable + gst_amt, 2)
    return subtotal, discount_amt, taxable, gst_amt, grand

def build_grouped_pdf_rows(df: pd.DataFrame):
    rows = [["Service", "Details", "Annual Fees (Rs.)"]]
    spans, styles = [], []
    r = 1
    for svc, grp in df.groupby("Service", sort=True):
        rows.append([svc, "", ""])
        spans.append(("SPAN", (0, r), (1, r)))
        styles.append(("BACKGROUND", (0, r), (-1, r), colors.HexColor("#f7f7f7")))
        styles.append(("FONTNAME", (0, r), (-1, r), "Helvetica-Bold"))
        styles.append(("ALIGN", (0, r), (-1, r), "LEFT"))
        r += 1
        for _, row in grp.iterrows():
            amt = pd.to_numeric(row["Annual Fees (Rs.)"], errors="coerce"); amt = 0.0 if pd.isna(amt) else float(amt)
            rows.append(["", row["Details"], money(amt)])
            r += 1
    return rows, spans, styles

def make_pdf(client_name, client_type, quote_no, df_quote, subtotal, discount_pct, discount_amt, gst_pct, gst_amt, grand, letterhead=False):
    import os
    from reportlab.platypus import Image
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen import canvas as canvas_mod

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4, leftMargin=18*mm, rightMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm
    )
    styles = getSampleStyleSheet()
    story = []

    def find_logo_path():
        for name in ("logo.png","logo.jpg","logo.jpeg"):
            if os.path.exists(name): return name
        return None

    logo_path = find_logo_path()
    if logo_path and not letterhead:
        ir = ImageReader(logo_path)
        ow, oh = ir.getSize(); max_w, max_h = 30*mm, 30*mm
        r = min(max_w/ow, max_h/oh); story.append(Image(logo_path, width=ow*r, height=oh*r)); story.append(Spacer(1, 4))

    story.append(Paragraph("<b>V. Purohit & Associates</b>", styles["Title"]))
    story.append(Paragraph("<b>Quotation</b>", styles["h2"])); story.append(Spacer(1, 6))
    meta_html = (
        f"<b>Quotation No.:</b> {quote_no}<br/>"
        f"<b>Client Name:</b> {client_name}<br/>"
        f"<b>Client Type:</b> {client_type}<br/>"
        f"<b>Date:</b> {datetime.now().strftime('%d-%b-%Y')}"
    )
    story.append(Paragraph(meta_html, styles["Normal"])); story.append(Spacer(1, 10))

    table_rows, spans, extra_styles = build_grouped_pdf_rows(df_quote)
    table = Table(table_rows, colWidths=[70*mm, 85*mm, 30*mm], repeatRows=1)
    base_style = [
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f0f0f0")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("VALIGN", (0,0), (-1,0), "MIDDLE"),
        ("BOTTOMPADDING", (0,0), (-1,0), 8),
        ("TOPPADDING", (0,0), (-1,0), 8),
        ("LINEBELOW", (0,0), (-1,0), 0.5, colors.grey),
        ("ALIGN", (2,1), (2,-1), "RIGHT"),
        ("INNERGRID", (0,1), (-1,-1), 0.3, colors.HexColor("#d9d9d9")),
    ]
    table.setStyle(TableStyle(base_style + spans + extra_styles))
    story.append(table); story.append(Spacer(1, 8))

    tot_lines = [
        ["", "Subtotal", money(subtotal)],
        ["", f"Discount ({discount_pct:.0f}%)", f"- {money(discount_amt)}"] if discount_amt > 0 else ["", "Discount (0%)", money(0)],
        ["", "Taxable Amount", money(subtotal - discount_amt)],
        ["", f"GST ({gst_pct:.0f}%)", money(gst_amt)],
        ["", "Grand Total", money(grand)],
    ]
    t2 = Table([["","", ""], *tot_lines], colWidths=[70*mm, 85*mm, 30*mm])
    t2.setStyle(TableStyle([
        ("ALIGN", (2,0), (2,-1), "RIGHT"),
        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
        ("BACKGROUND", (0,-1), (-1,-1), colors.HexColor("#fafafa")),
        ("LINEABOVE", (0,1), (-1,1), 0.5, colors.grey),
        ("BOTTOMPADDING", (0,1), (-1,-1), 6),
        ("TOPPADDING", (0,1), (-1,-1), 4),
    ]))
    story.append(t2); story.append(Spacer(1, 8))

    notes = (
        "<b>Note:</b><br/>"
        "1. The fees are exclusive of taxes and out-of-pocket expenses.<br/>"
        "2. GST as above.<br/>"
        "3. Our scope is limited to the services listed above.<br/>"
        "4. The above quotation is valid for a period of 30 days."
    )
    story.append(Paragraph(notes, styles["Normal"]))

    def _decorate(canv, doc_):
        logo_path_local = None
        try:
            for n in ("logo.png","logo.jpg","logo.jpeg"):
                import os
                if os.path.exists(n): logo_path_local = n; break
        except Exception: pass

        if st.session_state.get("letterhead", False) and logo_path_local:
            try:
                from reportlab.lib.utils import ImageReader
                canv.saveState()
                if hasattr(canv, "setFillAlpha"): canv.setFillAlpha(0.07)
                ir = ImageReader(logo_path_local)
                ow, oh = ir.getSize(); w = A4[0] - 60*mm; r = w / ow; h = oh * r
                x = 30*mm; y = (A4[1] - h)/2
                canv.drawImage(logo_path_local, x, y, width=w, height=h, preserveAspectRatio=True, mask="auto")
                canv.restoreState()
            except Exception: pass

        canv.saveState()
        canv.setFont("Helvetica", 8)
        line1 = "Office No. 5, Ground Floor, Adeshwar Arcade Commercial Premises CSL, Andheri - Kurla Road,"
        line2 = "Opp. Sangam Cinema, Andheri East, Mumbai - 400093.  Email: info@vpurohit.com, Contact: +91-8369508539"
        canv.drawCentredString(A4[0] / 2, 12*mm + 4, line1)
        canv.drawCentredString(A4[0] / 2, 12*mm - 6, line2)
        page = canv.getPageNumber(); canv.drawRightString(A4[0] - 18*mm, 12*mm - 18, f"Page {page}")
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
st.title(APP_TITLE); st.caption("Generate matrix-driven quotations and export to PDF")

with st.sidebar:
    st.subheader("Data")
    uploaded = st.file_uploader("Upload matrices.xlsx (optional)", type=["xlsx"])
    try:
        df_app, df_fees, source = load_matrices(uploaded)
        st.write(f"Source: **{source}**")
        st.write(f"Applicability rows: **{len(df_app):,}**")
        st.write(f"Fees rows: **{len(df_fees):,}**")

        st.subheader("Options")
        st.session_state["discount_pct"] = st.number_input("Discount %", 0, 100, int(st.session_state["discount_pct"]), 1)
        st.session_state["gst_pct"] = st.number_input("GST %", 0, 100, int(st.session_state["gst_pct"]), 1)
        st.session_state["letterhead"] = st.checkbox("Letterhead mode (watermark logo)", value=st.session_state["letterhead"])

        with st.expander("Data status", expanded=False):
            service_defs = len(df_app[["Service","SubService"]].drop_duplicates())
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
        st.error(f"Error loading matrices: {e}"); st.stop()

# ----- Form to generate initial table -----
with st.form("quote_form", clear_on_submit=False):
    client_name = st.text_input("Client Name*", st.session_state.get("client_name",""))
    client_type = st.selectbox("Client Type*", CLIENT_TYPES, index=0)

    ct_norm = normalize_str(client_type)
    app_ct = df_app[(df_app["ClientType"] == ct_norm) & (df_app["Applicable"] == True)]

    selected_accounting = st.radio("Accounting – choose one plan", ACCOUNTING_PLANS, index=3, horizontal=True)

    event_options = app_ct.loc[app_ct["Service"] == normalize_str(EVENT_SERVICE), "SubService"].dropna().unique().tolist()
    event_options_tc = sorted([title_with_acronyms(s) for s in event_options if s])
    selected_event_tc = st.multiselect(
        f"{title_with_acronyms(EVENT_SERVICE)} – select sub-services (choose any)",
        event_options_tc, default=[], help="Only selected items will be included in the quotation."
    )

    pt_options = app_ct.loc[app_ct["Service"] == normalize_str(PT_SERVICE), "SubService"].dropna().unique().tolist()
    pt_options_tc = sorted([title_with_acronyms(s) for s in pt_options if s])
    selected_pt_tc = st.radio(
        f"{title_with_acronyms(PT_SERVICE)} – choose one",
        pt_options_tc if pt_options_tc else ["(Not applicable)"], index=0, horizontal=True
    )
    if selected_pt_tc == "(Not applicable)": selected_pt_tc = None

    submit = st.form_submit_button("Generate Table")

# ----- Build grouped editor rows on submit -----
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
            # Create grouped rows with stable keys; keep input order
            rows = []
            for idx, r in df_quote.iterrows():
                svc, det, amt = r["Service"], r["Details"], float(pd.to_numeric(r["Annual Fees (Rs.)"], errors="coerce") or 0.0)
                rid = f"{svc}|{det}"
                rows.append({"id": rid, "service": svc, "details": det, "fee": amt, "base_fee": amt, "include": True})
            st.session_state["rows"] = rows
            st.session_state["client_name"] = client_name
            st.session_state["client_type"] = client_type
            st.session_state["editor_active"] = True

# ----- Grouped, editable UI (service headers centered, sub-rows indented) -----
def render_grouped_editor(rows: list[dict]):
    if not rows:
        st.info("No rows to display."); return pd.DataFrame()

    # Preserve service order by first appearance
    service_order = []
    seen = set()
    for r in rows:
        if r["service"] not in seen:
            seen.add(r["service"]); service_order.append(r["service"])

    # Per-service bulk actions + per-row controls
    for svc in service_order:
        svc_rows = [r for r in rows if r["service"] == svc]

        # Centered service header (merged look)
        st.markdown(
            f"""<div style="text-align:center; font-weight:700; padding:6px 0; background:#f7f7f7; border:1px solid #e6e6e6; border-radius:6px; margin-top:10px;">
            {svc}
            </div>""",
            unsafe_allow_html=True,
        )

        # Bulk include/exclude for this service
        b1, b2, _ = st.columns([0.18, 0.18, 0.64])
        if b1.button("Include all", key=f"inc_all_{hash(svc)}"):
            for r in svc_rows: st.session_state[f"inc_{r['id']}"] = True
        if b2.button("Exclude all", key=f"exc_all_{hash(svc)}"):
            for r in svc_rows: st.session_state[f"inc_{r['id']}"] = False

        # Sub-rows, indented
        for r in svc_rows:
            c0, c1, c2 = st.columns([0.08, 0.70, 0.22])
            inc_key = f"inc_{r['id']}"
            fee_key = f"fee_{r['id']}"
            # Defaults on first render
            if inc_key not in st.session_state: st.session_state[inc_key] = r["include"]
            if fee_key not in st.session_state: st.session_state[fee_key] = r["fee"]

            with c0:
                st.checkbox(" ", value=st.session_state[inc_key], key=inc_key, help="Include / exclude this row")

            with c1:
                st.markdown(f"""<div style="padding-left:16px;">{r['details']}</div>""", unsafe_allow_html=True)

            with c2:
                st.number_input(
                    "Annual Fees (Rs.)", min_value=0, step=100,
                    value=int(st.session_state[fee_key]) if isinstance(st.session_state[fee_key], (int, float)) else 0,
                    key=fee_key, label_visibility="collapsed",
                    help="Edit the fee; totals & PDF will use this value."
                )

    # Build DataFrame from current state
    data = []
    for r in rows:
        include = bool(st.session_state.get(f"inc_{r['id']}", True))
        fee_val = st.session_state.get(f"fee_{r['id']}", r["fee"])
        try:
            fee_val = float(fee_val)
        except Exception:
            fee_val = 0.0
        if include:
            data.append({"Service": r["service"], "Details": r["details"], "Annual Fees (Rs.)": fee_val})
    return pd.DataFrame(data)

# ----- Live editor + totals + PDF -----
if st.session_state["editor_active"] and st.session_state["rows"]:
    st.success("Quotation ready. Toggle include and edit fees below; totals update live.")
    df_selected = render_grouped_editor(st.session_state["rows"])

    subtotal, discount_amt, taxable, gst_amt, grand = compute_totals(
        df_selected, st.session_state["discount_pct"], st.session_state["gst_pct"]
    )

    c1, c2, c3, c4 = st.columns([1,1,1,1])
    with c1: st.write(f"**Subtotal (Rs.):** {money(subtotal)}")
    with c2: st.write(f"**Discount ({st.session_state['discount_pct']}%) (Rs.):** {money(discount_amt)}")
    with c3: st.write(f"**GST ({st.session_state['gst_pct']}%) (Rs.):** {money(gst_amt)}")
    with c4: st.write(f"**Grand Total (Rs.):** {money(grand)}")

    cA, cB, cC = st.columns([1,1,1])
    with cA:
        if st.button("Reset fees to matrix"):
            for r in st.session_state["rows"]:
                st.session_state[f"fee_{r['id']}"] = r["base_fee"]
            st.experimental_rerun()
    with cB:
        if st.button("Include all rows"):
            for r in st.session_state["rows"]:
                st.session_state[f"inc_{r['id']}"] = True
            st.experimental_rerun()
    with cC:
        if st.button("Start Over"):
            st.session_state["editor_active"] = False
            st.session_state["rows"] = []
            st.experimental_rerun()

    if df_selected.empty:
        st.info("All rows are excluded. Select at least one row to enable PDF.")
    else:
        # Build PDF (grouped format)
        pdf_bytes = make_pdf(
            st.session_state["client_name"], st.session_state["client_type"],
            datetime.now().strftime("QTN-%Y%m%d-%H%M%S"),
            df_selected, subtotal, float(st.session_state["discount_pct"]),
            discount_amt, float(st.session_state["gst_pct"]), gst_amt, grand,
            letterhead=st.session_state["letterhead"],
        )
        st.download_button(
            "⬇️ Download PDF",
            data=pdf_bytes,
            file_name=f"Quotation_{st.session_state['client_name'].replace(' ', '_')}.pdf",
            mime="application/pdf",
        )
