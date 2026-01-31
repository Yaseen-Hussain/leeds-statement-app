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
import zipfile
import io

# ---------------- HELPER FUNCTIONS ----------------
def parse_invoice_date(x):
    if pd.isna(x) or str(x).strip() == "":
        return pd.NaT

    # Excel serial date
    if isinstance(x, (int, float)):
        return pd.to_datetime(x, unit="D", origin="1899-12-30")

    # String date
    return pd.to_datetime(str(x).strip(), errors="coerce", dayfirst=True)

def format_amount(x):
    try:
        if pd.isna(x):
            return ""
        return f"{float(x):,.2f}"
    except (ValueError, TypeError):
        return ""

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


def generate_pdf(customer, today, opening_balance,
                 invoice_total, total_received,
                 closing_balance, date_range, rows):
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
        story.append(Spacer(1, 0.2*cm))

    company_style = ParagraphStyle(
        "CompanyName",
        parent=styles["Normal"],
        fontSize=16,
        leading=20,
        alignment=1,        # Centered
        fontName="Times-Bold"
    ) 

    story.append(Paragraph("Leeds Gifts Trading", company_style))
    story.append(Spacer(1, 0.2*cm))

    story.append(Paragraph(
        "<b>TRN:</b> 100465234100003<br/>"
        "<b>Email:</b> leedsgiftstrading123@gmail.com | "
        "<b>Mobile:</b> 0551423298",
        styles["Normal"]
    ))

    story.append(Spacer(1, 0.6*cm))



    # ---------- TITLES ----------
    title_style = ParagraphStyle(
        "Title",
        parent=styles["Normal"],
        fontSize=12,
        leading=18,
        alignment=1,
        fontName="Helvetica-Bold"
    )



    story.append(Paragraph("STATEMENT OF ACCOUNT", title_style))
    story.append(Spacer(1, 0.8*cm))

    # ---------- META ----------
    story.append(Paragraph(f"<b>Customer:</b> {customer}", styles["Normal"]))
    story.append(Paragraph(f"<b>Statement as of:</b> {today}", styles["Normal"]))
    story.append(Paragraph(f"<b>Statement period:</b> {date_range}", styles["Normal"]))
    story.append(Spacer(1, 0.6*cm))

    # ---------- SUMMARY BLOCK ----------
    summary_data = [
        ["Opening Balance", f"AED {opening_balance:,.2f}"],
        ["Invoices During Period", f"AED {invoice_total:,.2f}"],
        ["Payments Received", f"AED {total_received:,.2f}"],
        ["Closing Balance Outstanding", f"AED {closing_balance:,.2f}"],
    ]

    summary_table = Table(
        summary_data,
        colWidths=[9*cm, 5*cm]
    )

    summary_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("BACKGROUND", (0, -1), (-1, -1), colors.HexColor("#e6e6e6")),
    ]))

    story.append(summary_table)
    story.append(Spacer(1, 0.8*cm))






    # ---------- TABLE DATA ----------
    table_data = [
        [
        "S. No.",
        "Invoice Date",
        "Invoice No.",
        "Invoice Amount",
        "Amount Received",
        "Due Amount"
        ]
    ]

    for i, r in enumerate(rows, start=1):
        table_data.append([
            str(i),
            r["date"],
            r["inv"],
            r["invoice_amt"],
            r["received_amt"],
            r['due_amt']
        ])
    # ---- TOTAL ROW ----
    table_data.append([
        "Total", "", "",
        f"{invoice_total:,.2f}",
        f"{total_received:,.2f}",
        f"{closing_balance:,.2f}"
    ])



    table = Table(
        table_data,
        colWidths=[1.4*cm, 3*cm, 3.5*cm, 3*cm, 3*cm, 3*cm],
        repeatRows=1
    )

    table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f2a5a")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("ALIGN", (0, 0), (-1, 0), "CENTER"),

    # 👉 Right-align Amount columns (5 & 6)
    ("ALIGN", (3, 1), (5, -1), "RIGHT"),

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
        "This is a system-generated Statement of Account and does not require a signature.<br/>"
        "For discrepancies, contact: leedsgiftstrading123@gmail.com",
        ParagraphStyle("Footer", fontSize=8, alignment=1)
        )
    story.append(footer)

    doc.build(story)
    buffer.seek(0)
    return buffer

today = datetime.date.today().strftime("%d-%b-%Y")

# ---------------- CONFIG ----------------
LINES = {
    "Al Ain": "10-fOy3E-ni7XBtbC6zI-ivVrom481ZAZiS4xau0Ikpg",
    "Ajman": "1roBccgFgdXzGI3kEz5oRReADW7iKzWi-QboO2_scTKI",
    "Abu Dhabi": "10FTh4V5X8Y14u_6lIKUnaERlOX9b9F0x2xQ1vLXa2tY",
    "Fujairah": "1jD28UaXTLj9pTXrmtl17FfLoWM9qbUVnWudCG7oORQg",
    "Dubai & RAK": "1ZNW5OAeuCuVI9LNBtjlLWbwn6eVSqeT6kU-qzdhH-6I",
    "Sharjah": "1OA265FFNvGLvryTD8AZFpZPmXkpL155ce5BKFkO3LIg"
}

INVOICE_SHEET_NAME = "Invoice Wise"

# ---------------------------------------

st.set_page_config(
    page_title="Customer Statement",
    layout="centered",
    page_icon="logo.png"
)


# -------- PASSWORD GATE (CENTERED + ENTER SUPPORT) --------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:
        st.markdown(
            "<h2 style='text-align:center;'>LEEDS GIFTS TRADING</h2>",
            unsafe_allow_html=True
        )
        st.markdown(
            "<p style='text-align:center; color:grey;'>Customer Statement Portal</p>",
            unsafe_allow_html=True
        )

        with st.form("login_form"):
            pwd = st.text_input("Password", type="password")
            login_clicked = st.form_submit_button(
                "Login",
                use_container_width=True
            )

        if login_clicked:
            if pwd == st.secrets["APP_PASSWORD"]:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Wrong password")

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

    # ---- Invoice Amount ----
    df["Invoice Amount"] = (
        df["Invoice Amount"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )
    df["Invoice Amount"] = pd.to_numeric(df["Invoice Amount"], errors="coerce")


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
df["Invoice Date Parsed"] = df["Invoice Date"].apply(parse_invoice_date)

ENABLE_BULK_DOWNLOAD = True

# ================= BULK DOWNLOAD SECTION =================
if ENABLE_BULK_DOWNLOAD:
    st.markdown("Bulk Downloads (Line-wise)")

    confirm_bulk = st.checkbox(
    "I understand this will generate statements for all customers in this line (may take a minute)."
    )

    download_all_clicked = st.button(
        "Download all customer statements (ZIP)",
        disabled=not confirm_bulk
    )

    if download_all_clicked:
        st.info("Generating statements. Please wait...")

        progress = st.progress(0)

        zip_buffer = io.BytesIO()
        skipped_customers = 0

        customers_in_line = sorted(df["Customer Name"].dropna().unique())
        total_customers = len(customers_in_line)

        today_str = datetime.date.today().strftime("%Y-%m-%d")
        line_safe = line.replace(" ", "")

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for idx, cust in enumerate(customers_in_line, start=1):
                df_cust_bulk = df[df["Customer Name"] == cust].copy()

                if df_cust_bulk.empty:
                    skipped_customers += 1
                    continue

                df_cust_bulk["Invoice Date Parsed"] = df_cust_bulk["Invoice Date"].apply(parse_invoice_date)
                df_cust_bulk = df_cust_bulk.sort_values("Invoice Date Parsed")

                valid_dates = df_cust_bulk["Invoice Date Parsed"].dropna()
                if valid_dates.empty:
                    date_range = "All dates"
                else:
                    date_range = (
                        f"{valid_dates.iloc[0].strftime('%d-%b-%Y')} "
                        f"to {valid_dates.iloc[-1].strftime('%d-%b-%Y')}"
                    )

                total_due_bulk = pd.to_numeric(
                    df_cust_bulk["Due Amount"], errors="coerce"
                ).sum(skipna=True)

                if total_due_bulk <= 0:
                    skipped_customers += 1
                    progress.progress(idx / total_customers)
                    continue

                total_received_bulk = pd.to_numeric(
                    df_cust_bulk["Amount Received"], errors="coerce"
                ).sum(skipna=True)

                rows_bulk = []
                for _, r in df_cust_bulk.iterrows():
                    received_dt = parse_invoice_date(r["Received Date"])

                    rows_bulk.append({
                        "date": (
                            r["Invoice Date Parsed"].strftime("%d-%b-%Y")
                            if pd.notna(r["Invoice Date Parsed"]) else ""
                        ),
                        "inv": r["Invoice Number"],
                        "due_amt": format_amount(r["Due Amount"]),
                        "received_amt": format_amount(r["Amount Received"]),
                        "received_date": (
                            received_dt.strftime("%d-%b-%Y")
                            if pd.notna(received_dt) else ""
                        ),
                    })
                opening_balance = 0.0
                invoice_total = pd.to_numeric(
                    df_cust_bulk["Invoice Amount"], errors="coerce"
                ).sum(skipna=True)

                closing_balance = invoice_total - total_received_bulk
    
                pdf_buffer = generate_pdf(
                    customer=cust,
                    today=today,
                    opening_balance=opening_balance,
                    invoice_total=invoice_total,
                    total_received=total_received_bulk,
                    closing_balance=closing_balance,
                    date_range=date_range,
                    rows=rows_bulk
                )


                file_name = f"{cust}_{line_safe}_Statement_{today_str}.pdf"
                zipf.writestr(file_name, pdf_buffer.getvalue())

                progress.progress(idx / total_customers)

        zip_buffer.seek(0)

        st.success(
            f"ZIP ready. Generated {total_customers - skipped_customers} statements. "
            f"Skipped {skipped_customers} customers with no dues."
        )

        st.download_button(
            "Download ZIP",
            data=zip_buffer,
            file_name=f"{line_safe}_Statements_{today_str}.zip",
            mime="application/zip"
        )



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

#Date_range
    
st.markdown("### Statement Period")

period_type = st.radio(
    "Select statement type",
    ["All invoices", "Date range"],
    horizontal=True
)

start_date_input = None
end_date_input = None

if period_type == "Date range":
    col1, col2 = st.columns(2)

    with col1:
        start_date_input = st.date_input("Start date")

    with col2:
        end_date_input = st.date_input("End date")

    if start_date_input > end_date_input:
        st.error("Start date cannot be after End date")
        st.stop()

opening_balance = 0.0

if period_type == "Date range":
    start_dt = pd.to_datetime(start_date_input)

    df_before = df[
        (df["Customer Name"] == customer) &
        (df["Invoice Date Parsed"] < start_dt)
    ]

    opening_balance = pd.to_numeric(
        df_before["Due Amount"], errors="coerce"
    ).sum(skipna=True)


# ---------------- CUSTOMER FILTER ----------------
df_cust = df[df["Customer Name"] == customer].copy()

if df_cust.empty:
    st.warning("No invoices found for this customer.")
    st.stop()

df_cust["Invoice Date Parsed"] = df_cust["Invoice Date"].apply(parse_invoice_date)



# ---------------- SORT BY DATE ----------------
df_cust = df_cust.sort_values("Invoice Date Parsed")


# ---------------- DATE RANGE FILTER ----------------
if period_type == "Date range":
    start_dt = pd.to_datetime(start_date_input)
    end_dt = pd.to_datetime(end_date_input)

    df_cust = df_cust[
        (df_cust["Invoice Date Parsed"] >= start_dt) &
        (df_cust["Invoice Date Parsed"] <= end_dt)
    ]

    if df_cust.empty:
        st.warning("No invoices found in selected date range.")
        st.stop()



# ---------------- DATE RANGE (SAFE) ----------------
valid_dates = df_cust["Invoice Date Parsed"].dropna()

if valid_dates.empty:
    display_date_range = "All dates"
else:
    display_start = valid_dates.iloc[0].strftime("%d-%b-%Y")
    display_end = valid_dates.iloc[-1].strftime("%d-%b-%Y")
    display_date_range = f"{display_start} to {display_end}"



# ---------------- TOTALS ----------------
invoice_total = pd.to_numeric(
    df_cust["Invoice Amount"], errors="coerce"
).sum(skipna=True)

total_received = pd.to_numeric(
    df_cust["Amount Received"], errors="coerce"
).sum(skipna=True)

closing_balance = opening_balance + invoice_total - total_received



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

<div class="header">Statement of Account as on {{ date }}</div>
<p><b>Customer Name:</b> {{ customer }}</p>

<div class="summary">
Total outstanding amount: AED {{ total }}
</div>
<p><b>Date range:</b> {{ date_range }}</p>

<table>
<tr>
  <th>S. No.</th>
  <th>Invoice Date</th>
  <th>Invoice No.</th>
  <th>Due Amount</th>
</tr>

{% for row in rows %}
<tr>
  <td>{{ loop.index }}</td>
  <td>{{ row.date }}</td>
  <td>{{ row.inv }}</td>
  <td style="text-align:right;">{{ row.due_amt }}</td>
</tr>
{% endfor %}

<tr style="background-color:#e6e6e6; font-weight:bold;">
  <td colspan="3" style="text-align:center;">Total</td>
  <td style="text-align:right;">{{ total }}</td>
</tr>
</table>

</body>
</html>
"""




rows = []

for _, r in df_cust.iterrows():
    parsed_received_date = parse_invoice_date(r["Received Date"])

    rows.append({
    "date": (
        r["Invoice Date Parsed"].strftime("%d-%b-%Y")
        if pd.notna(r["Invoice Date Parsed"]) else ""
    ),
    "inv": r["Invoice Number"],
    "invoice_amt": format_amount(r["Invoice Amount"]),
    "received_amt": format_amount(r["Amount Received"]),
    "due_amt": format_amount(r["Due Amount"]),
})





html = Template(html_template).render(
    customer=customer,
    date=today,
    total=f"{closing_balance:,.2f}",
    date_range=display_date_range,
    rows=rows
)


st.markdown(html, unsafe_allow_html=True)




# -------- EXCEL --------
excel_df = df_cust[[
    "Invoice Date",
    "Invoice Number",
    "Due Amount",
    "Amount Received"
]]

excel_df.insert(0, "S. No.", range(1, len(excel_df) + 1))

output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    excel_df.to_excel(writer, index=False, startrow=3, sheet_name="Statement")

    ws = writer.sheets["Statement"]
    ws["A1"] = f"Customer: {customer}"
    ws["A2"] = f"Date range: {display_date_range}"




safe_range = display_date_range.replace(" ", "").replace("-", "")

pdf_buffer = generate_pdf(
    customer=customer,
    today=today,
    opening_balance=opening_balance,
    invoice_total=invoice_total,
    total_received=total_received,
    closing_balance=closing_balance,
    date_range=display_date_range,
    rows=rows
)



st.download_button(
    "Download PDF",
    data=pdf_buffer,
    file_name=f"{customer}_Statement_{safe_range}.pdf",
    mime="application/pdf"
)

st.download_button(
    "Download Excel",
    data=output.getvalue(),
    file_name=f"{customer}_Statement_{safe_range}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

