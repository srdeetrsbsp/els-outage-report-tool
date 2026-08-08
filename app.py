import io
import re
import urllib.parse
import pandas as pd
import requests
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

# Page Configuration
st.set_page_config(
    page_title="ELS Bilaspur - Daily Outage Tool", page_icon="🚆", layout="wide"
)

# Header Title
st.title("🚆 ELS Bilaspur - Daily Outage Report Generator")
st.markdown("---")

# Google Sheet Details
SHEET_ID = "1PPn0YqjTnwFm-xdG3jsmgiYfbSkUFce-cIGgmgVWkFg"


def load_google_sheet_tab(tab_name):
    """Fetch a tab from the connected Google Sheet master file."""
    encoded_tab = urllib.parse.quote(tab_name)
    url = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq?tqx=out:csv&sheet={encoded_tab}"
    try:
        return pd.read_csv(url)
    except Exception:
        return pd.DataFrame()


# Load Master Data from Google Sheet
with st.spinner("Connecting to Google Sheets Master Database..."):
    df_holding = load_google_sheet_tab("Loco_Holding")
    df_inside_shed = load_google_sheet_tab("Inside_Shed_List")
    df_targets = load_google_sheet_tab("Targets_and_Rules")

st.success("✅ Connected to Google Sheets Master Database!")

# Display Master Database Status
with st.expander("📊 View Master Database Status from Google Sheet"):
    col1, col2 = st.columns(2)
    with col1:
        st.write("**Loco Holding Count:**", len(df_holding))
        if not df_holding.empty:
            st.dataframe(df_holding.head(5), height=150)
    with col2:
        st.write("**Loco Inside Shed Count:**", len(df_inside_shed))
        if not df_inside_shed.empty:
            st.dataframe(df_inside_shed.head(5), height=150)

st.markdown("---")
st.header("1. Upload Daily FOIS / ICMS Files")

file_col1, file_col2 = st.columns(2)

with file_col1:
    uploaded_csv = st.file_uploader(
        "Upload Daily CSV File (e.g., LocoDepl.csv)", type=["csv"]
    )

with file_col2:
    uploaded_excel = st.file_uploader(
        "Upload Daily Excel File (e.g., Analysis.xlsx)", type=["xlsx", "xls"]
    )


def generate_pdf_report(
    report_date, loco_depl_df, inside_shed_locos, target_outage=180
):
    """Builds and returns a PDF file stream with full formatting."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=30,
        leftMargin=30,
        topMargin=30,
        bottomMargin=30,
    )
    styles = getSampleStyleSheet()

    # Custom Styles
    title_style = ParagraphStyle(
        "TitleStyle",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=13,
        alignment=1,
        spaceAfter=4,
    )
    subtitle_style = ParagraphStyle(
        "SubTitleStyle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        alignment=1,
        spaceAfter=10,
    )
    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=10,
        textColor=colors.HexColor("#003366"),
        spaceBefore=8,
        spaceAfter=4,
    )
    body_style = ParagraphStyle(
        "BodyTextCustom", parent=styles["Normal"], fontName="Helvetica", fontSize=9
    )
    table_cell_style = ParagraphStyle(
        "TableCell", parent=styles["Normal"], fontName="Helvetica", fontSize=8
    )
    table_header_style = ParagraphStyle(
        "TableHeader",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        textColor=colors.whitesmoke,
    )

    elements = []

    # Title & Header
    elements.append(Paragraph("SOUTH EAST CENTRAL RAILWAY", title_style))
    elements.append(
        Paragraph("ELECTRIC LOCO SHED, BILASPUR (द.पू.म.रे)", title_style)
    )
    elements.append(
        Paragraph(
            f"<b>DAILY LOCOMOTIVE OUTAGE & DETENTION REPORT</b> — Date: {report_date}",
            subtitle_style,
        )
    )
    elements.append(Spacer(1, 8))

    # Calculate statistics from uploaded data
    total_locos = len(loco_depl_df) if loco_depl_df is not None else 0

    # Filter out inside-shed locos
    if loco_depl_df is not None and not loco_depl_df.empty:
        col_loco = [c for c in loco_depl_df.columns if "LOCO" in c.upper()]
        loco_col = col_loco[0] if col_loco else loco_depl_df.columns[0]

        # Standardize loco numbers
        loco_depl_df["LOCO_CLEAN"] = (
            loco_depl_df[loco_col].astype(str).str.extract(r"(\d+)")[0]
        )
        traffic_outage_df = loco_depl_df[
            ~loco_depl_df["LOCO_CLEAN"].isin(inside_shed_locos)
        ]
        actual_outage = len(traffic_outage_df)
    else:
        actual_outage = 0
        traffic_outage_df = pd.DataFrame()

    loss_or_gain = actual_outage - target_outage

    # Summary Table Data
    summary_data = [
        [
            Paragraph("Target Outage", table_header_style),
            Paragraph("Actual Traffic Outage", table_header_style),
            Paragraph("Inside Shed Count", table_header_style),
            Paragraph("Outage Variance (Gain/Loss)", table_header_style),
        ],
        [
            Paragraph(f"<b>{target_outage}</b>", table_cell_style),
            Paragraph(f"<b>{actual_outage}</b>", table_cell_style),
            Paragraph(f"<b>{len(inside_shed_locos)}</b>", table_cell_style),
            Paragraph(
                f"<b>{'+' if loss_or_gain>=0 else ''}{loss_or_gain}</b>",
                table_cell_style,
            ),
        ],
    ]

    t_summary = Table(summary_data, colWidths=[130, 130, 130, 140])
    t_summary.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#f0f4f8")),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    elements.append(t_summary)
    elements.append(Spacer(1, 12))

    # Section 2: Traffic / Operating Detentions Breakdown
    elements.append(
        Paragraph("1. Traffic & Operating Detention Annexure", heading_style)
    )

    detention_rows = [
        [
            Paragraph("S.N.", table_header_style),
            Paragraph("Loco No.", table_header_style),
            Paragraph("Type", table_header_style),
            Paragraph("Location / Event", table_header_style),
            Paragraph("Load Name", table_header_style),
            Paragraph("Detention Hours", table_header_style),
        ]
    ]

    if not traffic_outage_df.empty:
        # Find relevant columns
        sttn_col = [c for c in traffic_outage_df.columns if "STTN" in c.upper()]
        event_col = [
            c for c in traffic_outage_df.columns if "EVENT" in c.upper()
        ]
        hrs_col = [c for c in traffic_outage_df.columns if "OUTAGE" in c.upper()]
        load_col = [c for c in traffic_outage_df.columns if "LOAD" in c.upper()]
        type_col = [c for c in traffic_outage_df.columns if "TYPE" in c.upper()]

        sttn_name = sttn_col[0] if sttn_col else ""
        event_name = event_col[0] if event_col else ""
        hrs_name = hrs_col[0] if hrs_col else ""
        load_name = load_col[0] if load_col else ""
        type_name = type_col[0] if type_col else ""

        count = 1
        for _, row in traffic_outage_df.head(25).iterrows():
            loco_num = str(row.get("LOCO_CLEAN", ""))
            l_type = str(row.get(type_name, "")) if type_name else ""
            sttn = str(row.get(sttn_name, "")) if sttn_name else ""
            evt = str(row.get(event_name, "")) if event_name else ""
            load = (
                str(row.get(load_name, "-"))
                if load_name and pd.notna(row.get(load_name))
                else "-"
            )
            hrs = (
                str(row.get(hrs_name, "-"))
                if hrs_name and pd.notna(row.get(hrs_name))
                else "-"
            )

            location_str = f"{sttn} ({evt})" if evt else sttn

            detention_rows.append(
                [
                    Paragraph(str(count), table_cell_style),
                    Paragraph(loco_num, table_cell_style),
                    Paragraph(l_type, table_cell_style),
                    Paragraph(location_str, table_cell_style),
                    Paragraph(load, table_cell_style),
                    Paragraph(hrs, table_cell_style),
                ]
            )
            count += 1
    else:
        detention_rows.append(
            [Paragraph("-", table_cell_style)] * 6
        )

    t_detention = Table(
        detention_rows, colWidths=[30, 70, 70, 160, 130, 70]
    )
    t_detention.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    elements.append(t_detention)
    elements.append(Spacer(1, 15))

    # Standard Railway Copy Distribution Footer
    elements.append(
        Paragraph("<b>Copy to:</b> DRM/BSP, ADRM/BSP, Sr.DEE/TRS/BSP for kind information please.", body_style)
    )

    doc.build(elements)
    buffer.seek(0)
    return buffer


if uploaded_csv is not None:
    # Read CSV
    try:
        csv_df = pd.read_csv(uploaded_csv, skiprows=2)
        st.success(f"Successfully processed CSV file: {len(csv_df)} records loaded.")
    except Exception:
        uploaded_csv.seek(0)
        csv_df = pd.read_csv(uploaded_csv)
        st.success(f"Successfully processed CSV file: {len(csv_df)} records loaded.")

    # Prepare Inside Shed Loco List from Google Sheet
    inside_shed_list = []
    if not df_inside_shed.empty:
        col_name = [c for c in df_inside_shed.columns if "LOCO" in c.upper()]
        if col_name:
            inside_shed_list = (
                df_inside_shed[col_name[0]].astype(str).str.extract(r"(\d+)")[0].tolist()
            )

    st.markdown("---")
    st.header("2. Generate Daily Outage PDF Report")

    report_date_input = st.date_input("Select Report Date")

    if st.button("🚀 Generate PDF Report"):
        pdf_data = generate_pdf_report(
            report_date=report_date_input.strftime("%d-%m-%Y"),
            loco_depl_df=csv_df,
            inside_shed_locos=inside_shed_list,
            target_outage=180,
        )

        st.success("🎉 PDF Outage Report generated successfully!")

        st.download_button(
            label="📄 Download PDF Outage Report",
            data=pdf_data,
            file_name=f"Daily_Outage_Report_ELS_BSP_{report_date_input.strftime('%Y%m%d')}.pdf",
            mime="application/pdf",
        )
else:
    st.info("👆 Please upload your daily CSV file above to generate the PDF report.")
