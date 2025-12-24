import streamlit as st
import pandas as pd
import datetime
import gspread
from google.oauth2.service_account import Credentials
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas


# ---------------- CONFIG ----------------
LINES = {
    "Al Ain": "1WmMFkyaqRhp2eAqx5m9LFgk6hM-HLNVfXjwhdh4TPa8",
}

INVOICE_SHEET_NAME = "Invoice Wise"

# ---------------------------------------

st.set_page_config(page_title="Customer Statement", layout="centered")

# -------- PASSWORD GATE --------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    pwd = st.text_input("Enter Password", type="password")
    if pwd == st.secrets["APP_PASSWORD"]:
        st.session_state.authenticated = True
        st.rerun()
    else:
        st.stop()

# -------- GOOGLE SHEETS --------
scope = ["https://www.googleapis.com/auth/spreadsheets.readonly"]
creds = Credentials.from_service_account_info(
    st.secrets["google"], scopes=scope
)
client = gspread.authorize(creds)

@st.cache_data(ttl=300)
def load_invoice_data(sheet_id, worksheet_name):
    sheet = client.open_by_key(sheet_id)
    ws = sheet.worksheet(worksheet_name)
    data = ws.get_all_records()
    df = pd.DataFrame(data)
    df["Due Amount"] = pd.to_numeric(df["Due Amount"], errors="coerce")
    return df


# -------- UI --------
st.image("logo.png", width=200)
st.markdown("<h2 style='text-align:center'>Statement</h2>", unsafe_allow_html=True)

line = st.selectbox("Select Line", list(LINES.keys()))

df = load_invoice_data(LINES[line], INVOICE_SHEET_NAME)
df = df[df["Due Amount"] > 0]


customers = sorted(df["Customer Name"].unique())
customer = st.selectbox("Customer Name", customers)

if df_cust.empty:
    st.warning("No outstanding invoices for this customer.")
    st.stop()


df_cust = df[df["Customer Name"] == customer].copy()
df_cust["Invoice Date"] = pd.to_datetime(df_cust["Invoice Date"])
df_cust = df_cust.sort_values("Invoice Date")

total_due = df_cust["Due Amount"].sum()
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
Total Outstanding Amount: AED {{ total }}
</div>

<table>
<tr>
<th>S. No.</th>
<th>Invoice Date</th>
<th>Invoice No.</th>
<th>Amount</th>
</tr>

{% for row in rows %}
<tr>
<td>{{ loop.index }}</td>
<td>{{ row.date }}</td>
<td>{{ row.inv }}</td>
<td>{{ row.amt }}</td>
</tr>
{% endfor %}
</table>

</body>
</html>
"""

rows = []
for _, r in df_cust.iterrows():
    rows.append({
        "date": r["Invoice Date"].strftime("%d-%b-%Y"),
        "inv": r["Invoice Number"],
        "amt": f"{r['Due Amount']:,.2f}"
    })

html = Template(html_template).render(
    customer=customer,
    date=today,
    total=f"{total_due:,.2f}",
    rows=rows
)

st.markdown(html, unsafe_allow_html=True)

# -------- PDF --------
from reportlab.lib.units import cm

def generate_pdf(customer, today, total_due, rows):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    y = height - 2 * cm

    c.setFont("Helvetica-Bold", 14)
    c.drawCentredString(width / 2, y, "Outstanding Invoice Statement")
    y -= 1 * cm

    c.setFont("Helvetica", 10)
    c.drawString(2 * cm, y, f"Customer Name: {customer}")
    y -= 0.6 * cm
    c.drawString(2 * cm, y, f"As on Date: {today}")
    y -= 1 * cm

    c.setFont("Helvetica-Bold", 10)
    c.drawString(2 * cm, y, f"Total Outstanding Amount: AED {total_due:,.2f}")
    y -= 1 * cm

    # Table header
    c.setFont("Helvetica-Bold", 9)
    c.drawString(2 * cm, y, "S.No")
    c.drawString(3.5 * cm, y, "Invoice Date")
    c.drawString(7 * cm, y, "Invoice No")
    c.drawString(13 * cm, y, "Amount")
    y -= 0.4 * cm

    c.setFont("Helvetica", 9)
    for i, r in enumerate(rows, start=1):
        if y < 2 * cm:
            c.showPage()
            y = height - 2 * cm
            c.setFont("Helvetica-Bold", 9)
            c.drawString(2 * cm, y, "S.No")
            c.drawString(3.5 * cm, y, "Invoice Date")
            c.drawString(7 * cm, y, "Invoice No")
            c.drawString(13 * cm, y, "Amount")
            y -= 0.4 * cm
            c.setFont("Helvetica", 9)


        c.drawString(2 * cm, y, str(i))
        c.drawString(3.5 * cm, y, r["date"])
        c.drawString(7 * cm, y, r["inv"])
        c.drawRightString(18 * cm, y, r["amt"])
        y -= 0.35 * cm

    c.save()
    buffer.seek(0)
    return buffer


# -------- EXCEL --------
excel_df = df_cust[["Invoice Date", "Invoice Number", "Due Amount"]]
excel_df.insert(0, "S. No.", range(1, len(excel_df) + 1))

output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    excel_df.to_excel(writer, index=False, sheet_name="Statement")

st.download_button(
    "Download Excel",
    data=output.getvalue(),
    file_name=f"{customer}_statement.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

pdf_buffer = generate_pdf(customer, today, total_due, rows)

st.download_button(
    "Download PDF",
    data=pdf_buffer,
    file_name=f"{customer}_statement.pdf",
    mime="application/pdf"
)


# -------- WHATSAPP --------
msg = f"Customer Statement – {customer}%0AOutstanding as on {today}: AED {total_due:,.2f}"
st.markdown(
    f"[Share via WhatsApp](https://wa.me/?text={msg})",
    unsafe_allow_html=True
)
