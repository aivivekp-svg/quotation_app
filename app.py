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
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Flowable

APP_TITLE = "Quotation Generator – V. Purohit & Associates"

CLIENT_TYPES = [
    "PRIVATE LIMITED", "PROPRIETORSHIP", "INDIVIDUAL", "LLP", "HUF",
    "SOCIETY", "PARTNERSHIP FIRM", "FOREIGN ENTITY", "AOP/ BOI", "TRUST",
]

ACCOUNTING_PLANS = ["Monthly Accounting", "Quarterly Accounting", "Half Yearly Accounting", "Annual Accounting"]

EVENT_SERVICE = "EVENT BASED FILING"
PT_SERVICE = "PROFESSION TAX RETURNS"

ACRONYMS = ["GST", "GSTR", "PTEC", "PTRC", "ADT", "ROC", "TDS", "AOC", "MGT", "26QB", "26QC"]
GST_RATE_FIXED = 18  # always 18%

# ---------- Session defaults ----------
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
_ss_set("sig_bytes", None)
_ss_set("theme_choice", "Light")
_ss_set("brand_color", "#0F4C81")

# ---------- Helpers ----------
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
    xl = pd.ExcelFile("matrices.xlsx")  # from repo
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

    # Split out Event Based Filing (no UI selection)
    main_app, event_app = split_main_vs_event(applicable)

    # Profession Tax Returns → choose one (in main set)
    is_pt = main_app["Service"].eq(normalize_str(PT_SERVICE))
    if selected_pt_sub is not None:
        sel_pt = normalize_str(selected_pt_sub)
        main_app = pd.concat(
            [main_app.loc[~is_pt], main_app.loc[is_pt & (main_app["SubService"] == sel_pt)]],
            ignore_index=True,
        )

    # Merge each with Fees and format labels
    def _merge_and_format(df_in: pd.DataFrame) -> pd.DataFrame:
        if df_in.empty:
            return pd.DataFrame(columns=["Service", "Details", "Annual Fees (Rs.)"])
        q = df_in.merge(df_fees, on=["Service", "SubService", "ClientType"], how="left", validate="1:1")
        q["FeeINR"] = pd.to_numeric(q["FeeINR"], errors="coerce").fillna(0.0)
        service_key_upper = q["Service"].copy()
        svc_pretty = service_key_upper.map(title_with_acronyms)
        q["Service"] = [service_display_override(raw, pretty) for raw, pretty in zip(service_key_upper.tolist(), svc_pretty.tolist())]
        q["SubService"] = q["SubService"].map(title_with_acronyms)
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

# ---------- Letterhead background (Style B) ----------
class PageBackgroundCard(Flowable):
    """
    Draws a soft page tint and a white content 'card' under all content.
    Placed as the very first flowable in the story; wrap returns (0,0)
    so it does not affect layout.
    """
    def __init__(self, tint_color="#F7F9FC", card_inset=10*mm, border_color=colors.HexColor("#E5E7EB")):
        super().__init__()
        self.tint_color = colors.HexColor(tint_color)
        self.card_inset = card_inset
        self.border_color = border_color

    def wrap(self, availWidth, availHeight):
        return (0, 0)

    def draw(self):
        canv = self.canv
        pw, ph = A4
        canv.saveState()
        # Page tint (light)
        canv.setFillColor(self.tint_color)
        canv.rect(0, 0, pw, ph, stroke=0, fill=1)
        # White content card with subtle border
        x = self.card_inset
        y = self.card_inset
        w = pw - 2 * self.card_inset
        h = ph - 2 * self.card_inset
        canv.setFillColor(colors.white)
        canv.setStrokeColor(self.border_color)
        canv.setLineWidth(1)
        canv.rect(x, y, w, h, stroke=1, fill=1)
        canv.restoreState()

# --- PDF table builders ---
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
            rows.append(["", row["Details"], money_inr(amt)])
            r += 1
    return rows, styles

def build_event_pdf_rows(df: pd.DataFrame):
    rows = [["Details", "Fees<br/>(Rs.)"]]
    for _, row in df.iterrows():
        amt = pd.to_numeric(row["Annual Fees (Rs.)"], errors="coerce")
        amt = 0.0 if pd.isna(amt) else float(amt)
        rows.append([row["Details"], money_inr(amt)])
    return rows

def make_pdf(client_name: str, client_type: str, quote_no: str,
             df_quote: pd.DataFrame, df_event: pd.DataFrame,
             subtotal: float, discount_pct: float, discount_amt: float, gst_amt: float, grand: float,
             letterhead: bool = False, addr: Optional[str] = None, email: Optional[str] = None,
             phone: Optional[str] = None, signature_bytes: Optional[bytes] = None):
    import os
    from reportlab.platypus import Image
    from reportlab.lib.utils import ImageReader

    buf = io.BytesIO()
    # Keep margins so content sits comfortably inside the white card
    doc = SimpleDocTemplate(
        buf, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm, topMargin=16 * mm, bottomMargin=16 * mm
    )
    styles = getSampleStyleSheet()
    head_center = ParagraphStyle("HeadCenter", parent=styles["Normal"], alignment=TA_CENTER, fontName="Helvetica-Bold")
    story = []

    # ---- Background card first (drawn under all content) ----
    story.append(PageBackgroundCard(tint_color="#F7F9FC", card_inset=10*mm, border_color=colors.HexColor("#E5E7EB")))

    # Small logo (watermark handled later if letterhead=True)
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

    # MAIN TABLE
    table_rows, extra_styles = build_grouped_pdf_rows(df_quote)
    table_rows[0] = [
        Paragraph("<b>Service</b>", head_center),
        Paragraph("<b>Details</b>", head_center),
        Paragraph("<b>Annual Fees</b><br/><b>(Rs.)</b>", head_center),
    ]
    col_widths = [60 * mm, 80 * mm, 30 * mm]  # fits safely within 174mm frame width
    table = Table(table_rows, colWidths=col_widths, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("FONTSIZE", (0, 0), (-1, 0), 10),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f0f0f0")),
                ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, 0), 6),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                ("BOX", (0, 0), (-1, 0), 0.9, colors.grey),
                ("INNERGRID", (0, 0), (-1, 0), 0.9, colors.grey),
                ("FONTSIZE", (0, 1), (-1, -1), 10),
                ("TOPPADDING", (0, 1), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
                ("ALIGN", (2, 1), (2, -1), "RIGHT"),
                ("INNERGRID", (0, 1), (-1, -1), 0.3, colors.HexColor("#d9d9d9")),
                ("BOX", (0, 0), (-1, -1), 1.0, colors.black),
            ]
            + extra_styles
        )
    )
    story.append(table)
    story.append(Spacer(1, 8))

    # TOTALS
    tot_lines = [
        ["", "Subtotal", money_inr(subtotal)],
        ["", f"Discount ({discount_pct:.0f}%)", f"- {money_inr(discount_amt)}"] if discount_amt > 0 else ["", "Discount (0%)", money_inr(0)],
        ["", "Taxable Amount", money_inr(subtotal - discount_amt)],
        ["", "GST (18%)", money_inr(gst_amt)],
        ["", "Grand Total", money_inr(grand)],
    ]
    t2 = Table([["", "", ""], *tot_lines], colWidths=col_widths)
    t2.setStyle(
        TableStyle(
            [
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("ALIGN", (2, 0), (2, -1), "RIGHT"),
                ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
                ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#fafafa")),
                ("LINEABOVE", (0, 1), (-1, 1), 0.5, colors.grey),
                ("TOPPADDING", (0, 1), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
                ("BOX", (0, 0), (-1, -1), 1.0, colors.black),
            ]
        )
    )
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

    # EVENT-BASED TABLE (separate)
    if not df_event.empty:
        story.append(Paragraph("<b>Event-based charges (as applicable, not included in annual fees)</b>", styles["Normal"]))
        story.append(Spacer(1, 4))
        ev_rows = build_event_pdf_rows(df_event)
        ev_rows[0] = [
            Paragraph("<b>Details</b>", head_center),
            Paragraph("<b>Fees</b><br/><b>(Rs.)</b>", head_center),
        ]
        ev = Table(ev_rows, colWidths=[140 * mm, 30 * mm], repeatRows=1)
        ev.setStyle(
            TableStyle(
                [
                    ("FONTSIZE", (0, 0), (-1, 0), 10),
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f0f0f0")),
                    ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
                    ("TOPPADDING", (0, 0), (-1, 0), 6),
                    ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                    ("BOX", (0, 0), (-1, 0), 0.9, colors.grey),
                    ("INNERGRID", (0, 0), (-1, 0), 0.9, colors.grey),
                    ("FONTSIZE", (0, 1), (-1, -1), 10),
                    ("TOPPADDING", (0, 1), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 1), (-1, -1), 4),
                    ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                    ("INNERGRID", (0, 1), (-1, -1), 0.3, colors.HexColor("#d9d9d9")),
                    ("BOX", (0, 0), (-1, -1), 1.0, colors.black),
                ]
            )
        )
        story.append(ev)
        story.append(Spacer(1, 10))

    # Footer / optional watermark / page numbers
    def _decorate(canv, doc_):
        # Optional watermark (letterhead mode)
        if letterhead:
            try:
                import os
                from reportlab.lib.utils import ImageReader
                for n in ("logo.png", "logo.jpg", "logo.jpeg"):
                    if os.path.exists(n):
                        canv.saveState()
                        if hasattr(canv, "setFillAlpha"):
                            canv.setFillAlpha(0.07)
                        ir = ImageReader(n)
                        ow, oh = ir.getSize()
                        w = A4[0] - 60 * mm
                        r = w / ow
                        h = oh * r
                        x = 30 * mm
                        y = (A4[1] - h) / 2
                        canv.drawImage(n, x, y, width=w, height=h, preserveAspectRatio=True, mask="auto")
                        canv.restoreState()
                        break
            except Exception:
                pass

        # Footer lines/text
        canv.saveState()
        canv.setFont("Helvetica", 8)
        y_line = 20 * mm
        canv.setLineWidth(0.7)
        canv.line(18 * mm, y_line, A4[0] - 18 * mm, y_line)
        line1 = "Office No. 5, Ground Floor, Adeshwar Arcade Commercial Premises CSL, Andheri - Kurla Road,"
        line2 = "Opp. Sangam Cinema, Andheri East, Mumbai - 400093.  Email: info@vpurohit.com, Contact: +91-8369508539"
        canv.drawCentredString(A4[0] / 2, 12 * mm + 4, line1)
        canv.drawCentredString(A4[0] / 2, 12 * mm - 6, line2)
        canv.drawRightString(A4[0] - 18 * mm, 12 * mm - 18, f"Page {canv.getPageNumber()}")
        canv.restoreState()

    doc.build(story, onFirstPage=_decorate, onLaterPages=_decorate)
    return buf.getvalue()

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

# ---------- UI ----------
st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="centered")

# Load matrices FIRST
try:
    df_app, df_fees, source = load_matrices()
except Exception as e:
    st.error(f"Error loading matrices: {e}")
    st.stop()

# ---------- SIDEBAR ----------
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
    st.subheader("Appearance")
    st.session_state["theme_choice"] = st.radio("Theme", ["Light", "Dark"], horizontal=True,
                                               index=0 if st.session_state["theme_choice"] == "Light" else 1)
    st.session_state["brand_color"] = st.color_picker("Brand color", st.session_state["brand_color"])

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

# ---------- THEME CSS (minimal & safe for icons) ----------
theme_choice = st.session_state["theme_choice"]
brand_color = st.session_state["brand_color"]

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/icon?family=Material+Icons');
@import url('https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:wght@300..700&display=swap');

.stButton > button, .stDownloadButton > button {{
  background: {brand_color} !important; color: #fff !important; border: 0; border-radius: 6px;
}}
.stApp {{ background: {"#ffffff" if theme_choice=="Light" else "#0f1117"} !important; }}
html, body {{ color: {"#111" if theme_choice=="Light" else "#e6e6e6"} !important; }}

[data-testid="stSidebarCollapseButton"] span,
[data-testid="stExpanderToggleIcon"] span,
.material-icons, .material-symbols-outlined,
[class*="material-icons"], [class*="material-symbols"] {{
  font-family: 'Material Symbols Outlined','Material Icons' !important;
  -webkit-font-feature-settings: 'liga';
  -webkit-font-smoothing: antialiased;
}}

h1 > a, h2 > a, h3 > a, h4 > a {{ display: none !important; }}
.label-lg {{ font-size: 1.05rem; font-weight: 700; margin: 6px 0 2px 0; }}
</style>
""", unsafe_allow_html=True)

# ---------- MAIN HEADER ----------
st.title("Quotation Generator")
st.subheader("V. Purohit & Associates")

# ---------- FORM ----------
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

# ---------- EDITOR / TABLES ----------
if submit:
    if not client_name.strip():
        st.error("Please enter Client Name.")
    else:
        st.session_state["quote_no"] = datetime.now().strftime("QTN-%Y%m%d-%H%M%S")
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
            st.session_state["quote_df"] = df_main
            st.session_state["event_df"] = event_df.copy()
            st.session_state["client_name"] = client_name
            st.session_state["client_type"] = client_type
            st.session_state["client_addr"] = addr
            st.session_state["client_email"] = email
            st.session_state["client_phone"] = phone
            st.session_state["editor_active"] = True

if st.session_state["editor_active"] and (
    not st.session_state["quote_df"].empty or not st.session_state["event_df"].empty
):
    st.success("Edit annual fees below. Event-based charges are listed separately and not included in totals.")

    # Main table editor (fees editable, Include toggles)
    if not st.session_state["quote_df"].empty:
        edited = st.data_editor(
            st.session_state["quote_df"],
            use_container_width=True,
            disabled=["Service", "Details"],
            column_config={
                "Include": st.column_config.CheckboxColumn(help="Uncheck to remove this row from the proposal/PDF."),
                "Annual Fees (Rs.)": st.column_config.NumberColumn(
                    "Annual Fees (Rs.)", min_value=0, step=100, format="%.0f",
                    help="Edit the fee; totals & PDF will use this value."
                ),
            },
            num_rows="fixed",
            key="quote_editor",
        )
        st.session_state["quote_df"] = edited
        filtered = edited[edited["Include"] == True].drop(columns=["Include"])
        filtered["Annual Fees (Rs.)"] = pd.to_numeric(filtered["Annual Fees (Rs.)"], errors="coerce").fillna(0.0)
    else:
        filtered = pd.DataFrame(columns=["Service", "Details", "Annual Fees (Rs.)"])

    # Event-based editor (fees editable; not included in totals)
    event_df = st.session_state["event_df"].copy()
    if not event_df.empty:
        st.subheader("Event-based charges (as applicable)")
        st.caption("These are not included in the annual fees totals.")
        event_edited = st.data_editor(
            event_df,
            use_container_width=True,
            disabled=["Service", "Details"],
            column_config={
                "Annual Fees (Rs.)": st.column_config.NumberColumn(
                    "Fees (Rs.)", min_value=0, step=100, format="%.0f",
                    help="Edit the fee; shown separately and not included in totals."
                ),
            },
            num_rows="fixed",
            key="event_editor",
        )
        event_edited["Annual Fees (Rs.)"] = pd.to_numeric(event_edited["Annual Fees (Rs.)"], errors="coerce").fillna(0.0)
        st.session_state["event_df"] = event_edited

    # Totals for main table (UI shown in INR format)
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

    # Downloads (Excel + PDF)
    colA, colB = st.columns([1, 1])
    if not filtered.empty:
        xlsx_buf = io.BytesIO()
        try:
            with pd.ExcelWriter(xlsx_buf, engine="openpyxl") as writer:
                filtered.to_excel(writer, index=False, sheet_name="Annual Fees")
            colA.download_button(
                "⬇️ Download Excel (Annual Fees)",
                data=xlsx_buf.getvalue(),
                file_name="annual_fees_proposal.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_xlsx_main",
            )
        except Exception:
            colA.caption(":grey[Excel export unavailable (missing engine).]")

    if not st.session_state["event_df"].empty:
        ev_xlsx = io.BytesIO()
        try:
            with pd.ExcelWriter(ev_xlsx, engine="openpyxl") as writer:
                st.session_state["event_df"].to_excel(writer, index=False, sheet_name="Event-based")
            colB.download_button(
                "⬇️ Download Excel (Event-based)",
                data=ev_xlsx.getvalue(),
                file_name="event_based_charges.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_xlsx_event",
            )
        except Exception:
            colB.caption(":grey[Excel export for event-based unavailable (missing engine).]")

    # PDF (includes event table + notes)
    pdf_bytes = make_pdf(
        st.session_state["client_name"],
        st.session_state["client_type"],
        st.session_state["quote_no"],
        filtered,
        st.session_state["event_df"],
        subtotal,
        float(st.session_state["discount_pct"]),
        discount_amt,
        gst_amt,
        grand,
        letterhead=st.session_state["letterhead"],
        addr=st.session_state.get("client_addr", ""),
        email=st.session_state.get("client_email", ""),
        phone=st.session_state.get("client_phone", ""),
        signature_bytes=st.session_state.get("sig_bytes"),
    )
    st.download_button(
        "⬇️ Download PDF",
        data=pdf_bytes,
        file_name=f"Annual_Fees_Proposal_{st.session_state['client_name'].replace(' ', '_')}.pdf",
        mime="application/pdf",
        key="dl_pdf",
    )
