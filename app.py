import streamlit as st
import pandas as pd
import datetime
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
from jinja2 import Template
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from io import BytesIO
import os
from pathlib import Path
from PIL import Image




# ---------------- CONFIG ----------------
LINES = {
    "Al Ain": "10-fOy3E-ni7XBtbC6zI-ivVrom481ZAZiS4xau0Ikpg",
    "Ajman": "1roBccgFgdXzGI3kEz5oRReADW7iKzWi-QboO2_scTKI",
    "Abu Dhabi": "10FTh4V5X8Y14u_6lIKUnaERlOX9b9F0x2xQ1vLXa2tY"
}

INVOICE_SHEET_NAME = "Invoice Wise"

# ---------------------------------------

st.set_page_config(page_title="Customer Statement", layout="centered")

# -------- PASSWORD GATE --------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
    st.session_state.pwd_error = False

if not st.session_state.authenticated:
    pwd = st.text_input("Enter Password", type="password")

    if pwd:
        if pwd == st.secrets["APP_PASSWORD"]:
            st.session_state.authenticated = True
            st.session_state.pwd_error = False
            st.rerun()
        else:
            st.session_state.pwd_error = True

    if st.session_state.pwd_error:
        st.error("Access denied: wrong password.")

    st.stop()


# -------- GOOGLE SHEETS --------
scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = Credentials.from_service_account_info(
    st.secrets["google"], scopes=scope
)
client = gspread.authorize(creds)

# @st.cache_data(ttl=300)
def load_invoice_data(sheet_id, worksheet_name):
    sheet = client.open_by_key(sheet_id)
    ws = sheet.worksheet(worksheet_name)

    values = ws.get_all_values(value_render_option="UNFORMATTED_VALUE")

    if len(values) < 3:
        return pd.DataFrame()

    headers = values[1]
    rows = values[2:]

    df = pd.DataFrame(rows, columns=headers)

    # ---- Due Amount ----
    df["Due Amount"] = (
        df["Due Amount"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df["Due Amount"] = pd.to_numeric(df["Due Amount"], errors="coerce")

    # ---- Amount Received ----
    df["Amount Received"] = (
        df["Amount Received"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df["Amount Received"] = pd.to_numeric(df["Amount Received"], errors="coerce")

    return df



# -------- UI --------
st.image("logo.png", width=200)
st.markdown("<h2 style='text-align:center'>Statement</h2>", unsafe_allow_html=True)

line = st.selectbox("Select Line", list(LINES.keys()))

df = load_invoice_data(LINES[line], INVOICE_SHEET_NAME)

customers = sorted(df["Customer Name"].unique())

customer = st.selectbox(
    "Customer Name",
    options=customers,
    index=None,
    placeholder="Type customer name"
)

if customer is None:
    st.info("Please select a customer to view the statement.")
    st.stop()



df_cust = df[df["Customer Name"] == customer].copy()

if df_cust.empty:
    st.warning("No outstanding invoices for this customer.")
    st.stop()

def parse_invoice_date(x):
    if pd.isna(x) or str(x).strip() == "":
        return pd.NaT

    # Excel serial date (number)
    if isinstance(x, (int, float)):
        return pd.to_datetime(x, unit="D", origin="1899-12-30")

    # String date (already formatted in sheet)
    return pd.to_datetime(x, errors="coerce")


df_cust["Invoice Date"] = df_cust["Invoice Date"].apply(parse_invoice_date)


df_cust = df_cust.sort_values("Invoice Date")

total_due = df_cust["Due Amount"].sum()
total_received = df_cust["Amount Received"].sum(skipna=True)
today = datetime.date.today().strftime("%d-%b-%Y")


# -------- HTML TEMPLATE --------
html_template = """
<html>
<head>
<style>
body { font-family: Arial; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid black; padding: 6px; text-align: center; }
th { background-color: #1f2a5a; color: white; }
.header { text-align: center; font-size: 18px; font-weight: bold; }
.summary { border: 1px solid black; padding: 8px; font-weight: bold; }
</style>
</head>
<body>

<div class="header">Outstanding Invoice Summary as on {{ date }}</div>
<p><b>Customer Name:</b> {{ customer }}</p>

<div class="summary">
Total outstanding amount: AED {{ total }}
</div>

<table>
<tr>
<th>S. No.</th>
<th>Invoice Date</th>
<th>Invoice No.</th>
<th>Due Amount</th>
<th>Amount Received</th>
<th>Received Date</th>
</tr>

{% for row in rows %}
<tr>
<td>{{ loop.index }}</td>
<td>{{ row.date }}</td>
<td>{{ row.inv }}</td>
<td style="text-align:right;">{{ row.amt }}</td>
<td style="text-align:right;">{{ row.received_amt }}</td>
<td>{{ row.received_date }}</td>
</tr>
{% endfor %}
<tr style="font-weight:bold; background-color:#e6e6e6;">
<td colspan="3" style="text-align:right;">Total</td>
<td style="text-align:right;">{{ total_due }}</td>
<td style="text-align:right;">{{ total_received }}</td>
<td></td>
</tr>
</table>

</body>
</html>
"""

def format_amount(x):
    try:
        if pd.isna(x):
            return ""
        return f"{float(x):,.2f}"
    except (ValueError, TypeError):
        return ""


rows = []

for _, r in df_cust.iterrows():
    parsed_received_date = parse_invoice_date(r["Received Date"])

    rows.append({
        "date": r["Invoice Date"].strftime("%d-%b-%Y"),
        "inv": r["Invoice Number"],
        "amt": format_amount(r["Due Amount"]),
        "received_amt": format_amount(r["Amount Received"]),
        "received_date": (
            parsed_received_date.strftime("%d-%b-%Y")
            if pd.notna(parsed_received_date) else ""
        )
    })



html = Template(html_template).render(
    customer=customer,
    date=today,
    total=f"{total_due:,.2f}",
    total_received=f"{total_received:,.2f}",
    rows=rows
)

st.markdown(html, unsafe_allow_html=True)

# -------- PDF --------
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from io import BytesIO
from pathlib import Path
from PIL import Image
from reportlab.platypus import Image


def generate_pdf(customer, today, total_due, rows):
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        rightMargin=2*cm,
        leftMargin=2*cm,
        topMargin=2*cm,
        bottomMargin=2*cm
    )

    styles = getSampleStyleSheet()
    story = []

    # ---------- LOGO ----------
    logo_path = Path(__file__).resolve().parent / "logo.png"

    if logo_path.exists():
        logo = Image(
            str(logo_path),
            width=4*cm,
            height=2*cm
        )
        logo.hAlign = "CENTER"

        story.append(logo)
        story.append(Spacer(1, 0.6*cm))


    # ---------- TITLES ----------
    title_style = ParagraphStyle(
        "Title",
        parent=styles["Normal"],
        fontSize=14,
        leading=18,
        alignment=1,
        fontName="Helvetica-Bold"
    )

    story.append(Paragraph("Outstanding Invoice Statement", title_style))
    story.append(Spacer(1, 0.8*cm))

    # ---------- META ----------
    story.append(Paragraph(f"<b>Customer Name:</b> {customer}", styles["Normal"]))
    story.append(Paragraph(f"<b>As on Date:</b> {today}", styles["Normal"]))
    story.append(Spacer(1, 0.4*cm))
    story.append(
        Paragraph(
            f"<b>Total outstanding amount:</b> AED {total_due:,.2f}",
            styles["Normal"]
        )
    )
    story.append(Spacer(1, 0.8*cm))

    # ---------- TABLE DATA ----------
    table_data = [
        [
        "S. No.",
        "Invoice Date",
        "Invoice No.",
        "Due Amount",
        "Amount Received",
        "Received Date"
        ]
    ]

    for i, r in enumerate(rows, start=1):
        table_data.append([
            str(i),
            r["date"],
            r["inv"],
            r["amt"],
            r["received_amt"],
            r["received_date"]
        ])
    # ---- TOTAL ROW ----
    table_data.append([
        "Total",
        "",
        "",
        f"{total_due:,.2f}",
        f"{total_received:,.2f}",
        ""
    ])


    table = Table(
        table_data,
        colWidths=[1.5*cm, 3.5*cm, 5*cm, 3*cm, 3*cm, 3.5*cm],
        repeatRows=1
    )

    table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f2a5a")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("ALIGN", (0, 0), (-1, 0), "CENTER"),

    # 👉 Right-align Amount columns (5 & 6)
    ("ALIGN", (3, 1), (4, -1), "RIGHT"),

    ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
    ("FONTSIZE", (0, 0), (-1, -1), 9),
    ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
    ("TOPPADDING", (0, 0), (-1, 0), 8),

    # Total row styling
    ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
    ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#e6e6e6")),
    ("ALIGN", (3, -1), (4, -1), "RIGHT"),
    ("SPAN", (0, -1), (2, -1)),
]))


    story.append(table)

    # ---------- FOOTER ----------
    story.append(Spacer(1, 0.8*cm))
    footer = Paragraph(
        "This is a system-generated statement and does not require a signature.",
        ParagraphStyle(
            "Footer",
            fontSize=8,
            alignment=1
        )
    )
    story.append(footer)

    doc.build(story)
    buffer.seek(0)
    return buffer


# -------- EXCEL --------
excel_df = df_cust[
    [
        "Invoice Date",
        "Invoice Number",
        "Due Amount",
        "Amount Received",
        "Received Date"
    ]
]
excel_df.insert(0, "S. No.", range(1, len(excel_df) + 1))

output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    excel_df.to_excel(writer, index=False, sheet_name="Statement")

pdf_buffer = generate_pdf(customer, today, total_due, rows)

st.download_button(
    "Download PDF",
    data=pdf_buffer,
    file_name=f"{customer}_Leedsgifts_Statement_{today}.pdf",
    mime="application/pdf"
)

st.download_button(
    "Download Excel",
    data=output.getvalue(),
    file_name=f"{customer}_Leedsgifts_Statement_{today}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

