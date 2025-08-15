import io
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
    "PRIVATE LIMITED",
    "PROPRIETORSHIP",
    "INDIVIDUAL",
    "LLP",
    "HUF",
    "SOCIETY",
    "PARTNERSHIP FIRM",
    "FOREIGN ENTITY",
    "AOP/ BOI",
    "TRUST",
]

ACCOUNTING_PLANS = [
    "Monthly Accounting",
    "Quarterly Accounting",
    "Half Yearly Accounting",
    "Annual Accounting",
]

# ------------------- Helpers -------------------
def normalize_str(x: str) -> str:
    return (x or "").strip().upper()

def load_matrices(uploaded_file: io.BytesIO | None):
    """Load Applicability and Fees matrices from uploaded file or local matrices.xlsx"""
    if uploaded_file is not None:
        xl = pd.ExcelFile(uploaded_file)
        source = "Uploaded file"
    else:
        xl = pd.ExcelFile("matrices.xlsx")
        source = "matrices.xlsx"

    df_app = xl.parse("Applicability").fillna("")
    df_fees = xl.parse("Fees").fillna("")

    # Normalize for robust joins
    for df in (df_app, df_fees):
        df["Service"] = df["Service"].map(normalize_str)
        df["SubService"] = df["SubService"].map(lambda v: normalize_str(v) if pd.notna(v) else "")
        df["ClientType"] = df["ClientType"].map(normalize_str)

    if "Applicable" in df_app.columns:
        df_app["Applicable"] = (
            df_app["Applicable"].astype(str).str.upper().isin(["TRUE", "1", "YES"])
        )
    if "FeeINR" in df_fees.columns:
        df_fees["FeeINR"] = pd.to_numeric(df_fees["FeeINR"], errors="coerce").fillna(0.0).astype(float)

    return df_app, df_fees, source

def build_quote(
    client_name: str,
    client_type: str,
    df_app: pd.DataFrame,
    df_fees: pd.DataFrame,
    selected_accounting: str | None = None,
):
    """Prepare the quotation table with requested formatting and filters."""
    ct = normalize_str(client_type)

    # Filter applicable services for client type
    applicable = (
        df_app
        .query("ClientType == @ct and Applicable == True")
        .loc[:, ["Service", "SubService", "ClientType"]]
        .copy()
    )

    # Keep only ONE Accounting plan (selected in UI)
    if selected_accounting:
        sel = normalize_str(selected_accounting)
        is_acc = applicable["Service"].str.upper().eq("ACCOUNTING")
        applicable = pd.concat(
            [
                applicable.loc[~is_acc],
                applicable.loc[is_acc & (applicable["SubService"].str.upper() == sel)],
            ],
            ignore_index=True,
        )

    # Merge with Fees
    quoted = applicable.merge(
        df_fees, on=["Service", "SubService", "ClientType"], how="left", validate="1:1"
    )

    # Clean presentation
    quoted["FeeINR"] = pd.to_numeric(quoted["FeeINR"], errors="coerce").fillna(0.0)
    quoted["Service"] = quoted["Service"].str.title()
    quoted["SubService"] = quoted["SubService"].str.title()
    quoted.sort_values(["Service", "SubService"], inplace=True)

    # Compute total
    total = float(quoted["FeeINR"].sum()) if not quoted.empty else 0.0

    # Final output:
    # - Drop ClientType column
    # - Rename SubService -> Details
    # - Rename FeeINR -> Annual Fees (Rs.)
    out = (
        quoted.drop(columns=["ClientType"], errors="ignore")
        .rename(columns={"SubService": "Details", "FeeINR": "Annual Fees (Rs.)"})
        .loc[:, ["Service", "Details", "Annual Fees (Rs.)"]]
    )

    return out, total

def make_pdf(client_name: str, client_type: str, df_quote: pd.DataFrame, total: float) -> bytes:
    """PDF with centered headers, right-aligned amount column, clean borders, notes, and optional logo."""
    import os
    from reportlab.platypus import Image
    from reportlab.lib.utils import ImageReader

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=A4,
        leftMargin=18*mm,
        rightMargin=18*mm,
        topMargin=15*mm,
        bottomMargin=15*mm,
    )

    styles = getSampleStyleSheet()
    story = []

    # --- Try to load a logo from repo root ---
    def find_logo_path():
        for name in ("logo.png", "logo.jpg", "logo.jpeg"):
            if os.path.exists(name):
                return name
        return None

    logo_path = find_logo_path()
    if logo_path:
        # Scale logo to a max box while preserving aspect ratio
        max_w, max_h = 30*mm, 30*mm
        ir = ImageReader(logo_path)
        orig_w, orig_h = ir.getSize()
        ratio = min(max_w / orig_w, max_h / orig_h)
        w, h = orig_w * ratio, orig_h * ratio
        story.append(Image(logo_path, width=w, height=h))
        story.append(Spacer(1, 4))

    # Header text
    story.append(Paragraph("<b>V. Purohit & Associates</b>", styles["Title"]))
    story.append(Paragraph("<b>Quotation</b>", styles["h2"]))
    story.append(Spacer(1, 6))

    meta_html = (
        f"<b>Client Name:</b> {client_name}<br/>"
        f"<b>Client Type:</b> {client_type}<br/>"
        f"<b>Date:</b> {datetime.now().strftime('%d-%b-%Y')}"
    )
    story.append(Paragraph(meta_html, styles["Normal"]))
    story.append(Spacer(1, 10))

    # --- Table ---
    headers = ["Service", "Details", "Annual Fees (Rs.)"]
    data = [headers]
    for _, row in df_quote.iterrows():
        data.append([row["Service"], row["Details"], f"{row['Annual Fees (Rs.)']:,.0f}"])
    data.append(["", "Total", f"{total:,.0f}"])

    table = Table(data, colWidths=[70*mm, 85*mm, 30*mm], repeatRows=1)
    table.setStyle(TableStyle([
        # Header
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#f0f0f0")),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("ALIGN", (0,0), (-1,0), "CENTER"),
        ("VALIGN", (0,0), (-1,0), "MIDDLE"),
        ("BOTTOMPADDING", (0,0), (-1,0), 8),
        ("TOPPADDING", (0,0), (-1,0), 8),
        ("LINEBELOW", (0,0), (-1,0), 0.5, colors.grey),
        # Body
        ("ALIGN", (2,1), (2,-1), "RIGHT"),
        ("INNERGRID", (0,1), (-1,-1), 0.3, colors.HexColor("#d9d9d9")),
        # Total row
        ("FONTNAME", (0,-1), (-1,-1), "Helvetica-Bold"),
        ("BACKGROUND", (0,-1), (-1,-1), colors.HexColor("#fafafa")),
    ]))
    story.append(table)
    story.append(Spacer(1, 8))

    # Notes block
    notes = (
        "<b>Note:</b><br/>"
        "1. The fees are exclusive of taxes and out-of-pocket expenses.<br/>"
        "2. GST 18% extra.<br/>"
        "3. Our scope is limited to the services listed above.<br/>"
        "4. The above quotation is valid for a period of 30 days."
    )
    story.append(Paragraph(notes, styles["Normal"]))

    doc.build(story)
    return buf.getvalue()

def build_status(df_app: pd.DataFrame, df_fees: pd.DataFrame) -> pd.DataFrame:
    """Return a status table with Applicable counts and Missing/Zero fee counts per ClientType."""
    active = df_app[df_app["Applicable"] == True].copy()
    # Counts of applicable rows
    counts = (
        active.groupby("ClientType").size().reindex(CLIENT_TYPES, fill_value=0).reset_index(name="Applicable services")
    )
    # Missing or zero fees among applicable
    merged = active.merge(df_fees, on=["Service","SubService","ClientType"], how="left")
    missing_mask = merged["FeeINR"].isna() | (pd.to_numeric(merged["FeeINR"], errors="coerce").fillna(0.0) <= 0)
    miss = (
        merged[missing_mask]
        .groupby("ClientType").size()
        .reindex(CLIENT_TYPES, fill_value=0)
        .reset_index(name="Missing/Zero fees")
    )
    status = counts.merge(miss, on="ClientType")
    return status

# ------------------- UI -------------------
st.set_page_config(page_title=APP_TITLE, page_icon="📄", layout="centered")
st.title(APP_TITLE)
st.caption("Generate matrix-driven quotations and export to PDF")

with st.sidebar:
    st.subheader("Data")
    uploaded = st.file_uploader("Upload matrices.xlsx (optional)", type=["xlsx"])
    try:
        df_app, df_fees, source = load_matrices(uploaded)
        st.write(f"Source: **{source}**")
        st.write(f"Applicability rows: **{len(df_app):,}**")
        st.write(f"Fees rows: **{len(df_fees):,}**")
        with st.expander("Data status", expanded=False):
            service_defs = len(df_app[["Service","SubService"]].drop_duplicates())
            st.write(f"Service definitions: **{service_defs}**")
            status_df = build_status(df_app, df_fees)
            st.dataframe(status_df, use_container_width=True)
    except Exception as e:
        st.error(f"Error loading matrices: {e}")
        st.stop()

with st.form("quote_form", clear_on_submit=False):
    client_name = st.text_input("Client Name*", "")
    client_type = st.selectbox("Client Type*", CLIENT_TYPES, index=0)

    # Choose one accounting plan
    selected_accounting = st.radio(
        "Accounting – choose one plan",
        ACCOUNTING_PLANS,
        index=3,
        horizontal=True,
    )

    generate = st.form_submit_button("Generate Table")

if generate:
    if not client_name.strip():
        st.error("Please enter Client Name.")
    else:
        df_quote, total = build_quote(
            client_name,
            client_type,
            df_app,
            df_fees,
            selected_accounting=selected_accounting,
        )

        if df_quote.empty:
            st.warning("No applicable services found for the selected Client Type.")
        else:
            st.success("Quotation ready.")
            st.dataframe(df_quote, use_container_width=True)
            st.write(f"**Grand Total (Rs.):** {total:,.0f}")

            pdf_bytes = make_pdf(client_name, client_type, df_quote, total)
            st.download_button(
                "⬇️ Download PDF",
                data=pdf_bytes,
                file_name=f"Quotation_{client_name.replace(' ', '_')}.pdf",
                mime="application/pdf",
            )
