"""
Streamlit Billing System — Mobile Phone Business (PDF Invoice)
Run: streamlit run app.py
"""

import streamlit as st
import smtplib
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import re
import io
from datetime import datetime, date
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from xhtml2pdf import pisa

# ── CONFIG ──
SENDER_EMAIL = "smartchoicemobileskandy@gmail.com"
APP_PASSWORD = "xatoamcihdbrqzvy"
BUSINESS_NAME = "Smart Choice Mobiles Kandy"
BUSINESS_TAGLINE = "Your Trusted Mobile Partner"

st.set_page_config(page_title=f"{BUSINESS_NAME} Billing", page_icon="📱", layout="wide")

# Establish Google Sheets connection
conn = st.connection("gsheets", type=GSheetsConnection)

# ── CUSTOM CSS ──
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.main-header {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    padding: 2rem 2.5rem; border-radius: 14px; margin-bottom: 1.5rem; text-align: center;
}
.main-header h1 { color: #fff; font-size: 2.2rem; font-weight: 800; letter-spacing: 4px; margin: 0; }
.main-header p { color: #8892b0; font-size: 0.85rem; letter-spacing: 2px; margin: 4px 0 0 0; }
.section-label {
    font-size: 0.75rem; font-weight: 700; letter-spacing: 2px; text-transform: uppercase;
    color: #888; margin-bottom: 0.5rem; padding-bottom: 0.4rem; border-bottom: 2px solid #eee;
}
div[data-testid="stForm"] {
    border: 1px solid #e0e0e0; border-radius: 12px; padding: 1.5rem;
    background: #fafbfc; box-shadow: 0 2px 12px rgba(0,0,0,0.04);
}
.stButton > button[kind="formSubmit"],
div[data-testid="stForm"] button[kind="formSubmit"] {
    background: linear-gradient(135deg, #43a047, #2e7d32) !important;
    color: white !important; font-weight: 700 !important; letter-spacing: 1px !important;
    border: none !important; border-radius: 8px !important; padding: 0.6rem 2rem !important;
    font-size: 1rem !important; width: 100% !important;
}
div[data-testid="stForm"] button[kind="formSubmit"]:hover {
    background: linear-gradient(135deg, #388e3c, #1b5e20) !important;
}
</style>
""", unsafe_allow_html=True)

# ── HEADER ──
st.markdown(f"""
<div class="main-header">
    <h1>📱 {BUSINESS_NAME.upper()}</h1>
    <p>{BUSINESS_TAGLINE}</p>
</div>
""", unsafe_allow_html=True)


# ══════════════════════════════════════
#  HELPER FUNCTIONS
# ══════════════════════════════════════

def validate_email(email: str) -> bool:
    return bool(re.match(r"^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$", email))

def validate_imei(imei: str) -> bool:
    return bool(re.match(r"^\d{15}$", imei))




def build_pdf_html(d: dict) -> str:
    """Build a Deep Red/Charcoal/White corporate PDF invoice (xhtml2pdf-compatible)."""
    # Conditional Notes / Remarks section
    notes_section = ""
    if d.get("other_notes", "").strip():
        notes_section = f"""<table width="100%" cellpadding="0" cellspacing="0" style="margin-top:3px;">
    <tr>
      <td style="padding:8px 10px; background-color:#ffffff; border:1px solid #eeeeee;">
        <p style="margin:0 0 2px 0; font-size:11px; font-weight:bold; color:#222222; text-transform:uppercase; letter-spacing:1.5px;">Notes / Remarks</p>
        <p style="margin:0; font-size:11px; color:#444444; line-height:1.4;">{d['other_notes']}</p>
      </td>
    </tr>
  </table>"""
    exchange_rows = ""
    if d["payment_type"] == "Exchange":
        exchange_rows = f"""
      <tr>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#444; font-weight:bold; background-color:#f9f9f9; font-size:13px;">Old Phone Model</td>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#222222; background-color:#f9f9f9; font-size:13px;">{d['old_phone_model']}</td>
      </tr>
      <tr>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#444; font-weight:bold; background-color:#ffffff; font-size:13px;">Old Phone IMEI</td>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#222222; font-family:Courier; letter-spacing:1px; background-color:#ffffff; font-size:13px;">{d['old_phone_imei']}</td>
      </tr>"""

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
    @page {{ size: A4; margin: 8mm; background: #ffffff; }}
    body {{ font-family: Helvetica, Arial, sans-serif; color: #222222; margin: 0; padding: 0; }}
    p {{ margin: 0; }}
    .table-main {{ width: 100%; border-collapse: collapse; margin-top: 0; }}
    .table-main th {{ background-color: #990000; color: #ffffff; text-align: left; padding: 8px 12px; font-size: 13px; font-weight: bold; text-transform: uppercase; }}
    .table-main td {{ padding: 8px 12px; border-bottom: 1px solid #eeeeee; font-size: 13px; color: #222222; }}
    .table-totals {{ width: 100%; border-collapse: collapse; margin-top: 0; }}
</style>
</head>
<body>

    <table width="100%" cellpadding="8" style="background-color: #222222; border-bottom: 5px solid #990000;">
        <tr>
            <td align="center">
                <p style="color: #ffffff; font-size: 24px; font-weight: bold; letter-spacing: 2px; margin: 0 0 2px 0;">SMART CHOICE MOBILES</p>
                <p style="color: #cccccc; font-size: 10.5px; margin: 0;">Maligathanna, Gurudeniya &nbsp;&nbsp;|&nbsp;&nbsp; Phone / WhatsApp: 0723175373 &nbsp;&nbsp;|&nbsp;&nbsp; smartchoicemobileskandy@gmail.com</p>
            </td>
        </tr>
    </table>

    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top:30px;">
        <tr>
            <td width="50%" valign="top">
                <p style="color: #990000; font-size: 12px; font-weight: bold; text-transform: uppercase; margin: 0 0 4px 0;">Billed To</p>
                <p style="font-size: 13px; font-weight: bold; color: #222222; margin: 0 0 2px 0;">{d['customer_name']}</p>
                <p style="font-size: 11px; color: #444444; margin: 0 0 1px 0;">{d.get('customer_address', '')}</p>
                <p style="font-size: 11px; color: #444444; margin: 0 0 1px 0;">{d.get('customer_contact', '')}</p>
                <p style="font-size: 11px; color: #444444; margin: 0;">{d['customer_email']}</p>
            </td>

            <td width="50%" valign="top" align="right">
                <p style="color: #990000; font-size: 12px; font-weight: bold; text-transform: uppercase; margin: 0 0 4px 0; text-align: right;">Invoice Details</p>
                <table width="100%" cellpadding="2" cellspacing="0" border="0">
                    <tr>
                        <td width="50%" align="right" style="font-size: 11px; font-weight: bold; color: #222222;">Invoice #:</td>
                        <td width="50%" align="right" style="font-size: 11px; color: #444444;">{d['invoice_number']}</td>
                    </tr>
                    <tr>
                        <td width="50%" align="right" style="font-size: 11px; font-weight: bold; color: #222222;">Date:</td>
                        <td width="50%" align="right" style="font-size: 11px; color: #444444;">{d['date']}</td>
                    </tr>
                    <tr>
                        <td width="50%" align="right" style="font-size: 11px; font-weight: bold; color: #222222;">Payment:</td>
                        <td width="50%" align="right" style="font-size: 11px; color: #444444;">{d['payment_type']}</td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>

    <!-- ════════════ TRANSACTION TABLE ════════════ -->
    <table class="table-main" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #eeeeee; border-collapse:collapse; margin-top:25px;">
      <tr>
        <th style="background-color:#990000; padding:8px 12px; font-size:13px; font-weight:bold; color:#ffffff; text-transform:uppercase; width:40%;">Description</th>
        <th style="background-color:#990000; padding:8px 12px; font-size:13px; font-weight:bold; color:#ffffff; text-transform:uppercase;">Details</th>
      </tr>
      <tr>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#444; font-weight:bold; background-color:#ffffff; font-size:13px; width:40%;">New Phone Model</td>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#222222; font-weight:bold; background-color:#ffffff; font-size:13px;">{d['new_phone_model']}</td>
      </tr>
      <tr>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#444; font-weight:bold; background-color:#f9f9f9; font-size:13px; width:40%;">New Phone IMEI</td>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#222222; font-family:Courier; letter-spacing:1px; background-color:#f9f9f9; font-size:13px;">{d['new_phone_imei']}</td>
      </tr>
      <tr>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#444; font-weight:bold; background-color:#ffffff; font-size:13px; width:40%;">Warranty Period</td>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#222222; background-color:#ffffff; font-size:13px;">{d['warranty']}</td>
      </tr>
      <tr>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#444; font-weight:bold; background-color:#f9f9f9; font-size:13px; width:40%;">Payment Type</td>
        <td style="padding:8px 12px; border-bottom:1px solid #eeeeee; color:#222222; background-color:#f9f9f9; font-size:13px;">{d['payment_type']}</td>
      </tr>
      {exchange_rows}
    </table>

    <!-- ════════════ TOTALS ════════════ -->
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:5px;">
      <tr>
        <td width="55%"></td>
        <td width="45%">
          <table class="table-totals" width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
            <tr>
              <td style="padding:6px 12px; border-bottom:1px solid #eeeeee; font-size:12px; font-weight:bold; color:#666; text-transform:uppercase; letter-spacing:1px;">Total Amount</td>
              <td align="right" style="padding:6px 12px; border-bottom:1px solid #eeeeee; font-size:16px; font-weight:bold; color:#222222;">Rs. {d['total_price']}</td>
            </tr>
            <tr>
              <td style="padding:6px 12px; font-size:12px; font-weight:bold; color:#222222; text-transform:uppercase; letter-spacing:1px;">Balance Due</td>
              <td align="right" style="padding:6px 12px; font-size:18px; font-weight:bold; color:#990000;">Rs. {d['balance']}</td>
            </tr>
          </table>
        </td>
      </tr>
    </table>

    {notes_section}

    <!-- ════════════ TERMS & WARRANTY POLICY ════════════ -->
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:5px;">
      <tr>
        <td style="padding:6px 10px; background-color:#f9f9f9; border:1px solid #eeeeee;">
          <p style="margin:0 0 4px 0; font-size:11px; font-weight:bold; color:#222222; text-transform:uppercase; letter-spacing:1.5px;">Terms &amp; Warranty Policy</p>
          <p style="margin:0 0 2px 0; font-size:10px; color:#444444; line-height:1.4;">1. Warranty claims are valid ONLY for issues that occur strictly within the specified warranty period mentioned in this invoice.</p>
          <p style="margin:0 0 2px 0; font-size:10px; color:#444444; line-height:1.4;">2. The warranty seal/sticker applied on the device MUST be intact. Any tampering, tearing, or removal of the seal will immediately void the entire warranty.</p>
          <p style="margin:0 0 2px 0; font-size:10px; color:#444444; line-height:1.4;">3. The original invoice must be presented for any warranty claims.</p>
          <p style="margin:0 0 2px 0; font-size:10px; color:#444444; line-height:1.4;">4. Physical damage, liquid damage, and display drops are completely not covered under warranty.</p>
          <p style="margin:0; font-size:10px; color:#444444; line-height:1.4;">5. Software issues are subject to the standard company policy.</p>
        </td>
      </tr>
    </table>

    <!-- ════════════ FOOTER ════════════ -->
    <table width="100%" cellpadding="0" cellspacing="0" style="margin-top:5px;">
      <tr><td style="background-color:#990000; height:2px; font-size:1px;">&nbsp;</td></tr>
    </table>
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center" style="padding:4px 0 0 0;">
          <p style="margin:0 0 1px 0; font-size:9px; color:#555;">Thank you for choosing <b style="color:#222222;">Smart Choice Mobiles Kandy</b>!</p>
          <p style="margin:0; font-size:9px; color:#999;">Computer-generated document. No signature required.</p>
        </td>
      </tr>
    </table>

</body>
</html>"""


def generate_pdf(html: str) -> bytes:
    """Convert HTML to PDF bytes using xhtml2pdf."""
    buffer = io.BytesIO()
    pisa_status = pisa.CreatePDF(io.StringIO(html), dest=buffer)
    if pisa_status.err:
        raise RuntimeError(f"PDF generation failed with {pisa_status.err} error(s)")
    return buffer.getvalue()


def send_email_with_pdf(to: str, subject: str, pdf_bytes: bytes, inv_num: str) -> tuple[bool, str]:
    """Send a plain-text email with the PDF invoice attached."""
    msg = MIMEMultipart("mixed")
    msg["From"] = f"{BUSINESS_NAME} <{SENDER_EMAIL}>"
    msg["To"] = to
    msg["Subject"] = subject

    # Simple professional email body
    body = f"""Dear Customer,

Thank you for your purchase from {BUSINESS_NAME}.

Please find your official invoice ({inv_num}) attached as a PDF document.

If you have any questions regarding this invoice, please don't hesitate to contact us.

Best regards,
{BUSINESS_NAME}
{BUSINESS_TAGLINE}"""

    msg.attach(MIMEText(body, "plain"))

    # Attach PDF
    pdf_attachment = MIMEApplication(pdf_bytes, _subtype="pdf")
    pdf_attachment.add_header("Content-Disposition", "attachment", filename=f"{inv_num}.pdf")
    msg.attach(pdf_attachment)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(SENDER_EMAIL, APP_PASSWORD)
            s.sendmail(SENDER_EMAIL, to, msg.as_string())
        return True, ""
    except smtplib.SMTPAuthenticationError:
        return False, "Authentication failed. Check SENDER_EMAIL & APP_PASSWORD."
    except smtplib.SMTPRecipientsRefused:
        return False, f"Recipient '{to}' was refused by the server."
    except Exception as e:
        return False, str(e)


# ══════════════════════════════════════
#  PAYMENT TYPE (outside form for reactivity)
# ══════════════════════════════════════
st.markdown('<p class="section-label">💳 Payment Type</p>', unsafe_allow_html=True)
payment_type = st.radio("Payment Type", ["Full Cash", "Exchange"], horizontal=True, label_visibility="collapsed")

# ══════════════════════════════════════
#  STREAMLIT FORM
# ══════════════════════════════════════
with st.form("billing_form", clear_on_submit=False):

    # ── Row 1: Date & Customer ──
    st.markdown('<p class="section-label">📋 Customer & Date</p>', unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 1.5, 1.5])
    with c1:
        bill_date = st.date_input("Date", value=date.today())
    with c2:
        customer_name = st.text_input("Customer Name")
    with c3:
        customer_email = st.text_input("Customer Email")

    ca1, ca2 = st.columns(2)
    with ca1:
        customer_address = st.text_input("Customer Address")
    with ca2:
        customer_contact = st.text_input("Customer Contact No")

    st.divider()

    # ── Row 2: New Phone ──
    st.markdown('<p class="section-label">📦 New Phone Details</p>', unsafe_allow_html=True)
    c4, c5, c6 = st.columns(3)
    with c4:
        new_model = st.text_input("New Phone Model")
    with c5:
        new_imei = st.text_input("New Phone IMEI", max_chars=15, help="Must be exactly 15 digits")
    with c6:
        warranty = st.selectbox("Warranty Period", ["1 Month", "3 Months", "6 Months", "1 Year", "2 Years"])

    # ── Row 3: Old Phone (conditional — reactive via outside radio) ──
    old_model, old_imei = "", ""
    if payment_type == "Exchange":
        st.divider()
        st.markdown('<p class="section-label">🔄 Exchange — Old Phone</p>', unsafe_allow_html=True)
        cx, cy = st.columns(2)
        with cx:
            old_model = st.text_input("Old Phone Model")
        with cy:
            old_imei = st.text_input("Old Phone IMEI", max_chars=15, key="old_imei")

    st.divider()

    # ── Row 4: Pricing ──
    st.markdown('<p class="section-label">💰 Pricing</p>', unsafe_allow_html=True)
    p1, p2 = st.columns(2)
    with p1:
        total_price = st.text_input("Total Agreed Price (Rs.)", placeholder="e.g. 125000")
    with p2:
        balance = st.text_input("Balance Amount (Rs.)", placeholder="e.g. 50000")

    st.divider()

    # ── Other Notes ──
    st.markdown('<p class="section-label">📝 Additional Notes</p>', unsafe_allow_html=True)
    other_notes = st.text_area("Other Notes (Optional)", height=80, placeholder="Any additional remarks for the invoice…")

    st.markdown("")
    submitted = st.form_submit_button("⚡  Generate & Send Invoice")

# ── PROCESS SUBMISSION ──
if submitted:
    errors = []
    if not customer_name.strip():
        errors.append("Customer Name is required.")
    if not customer_email.strip() or not validate_email(customer_email):
        errors.append("A valid Customer Email is required.")
    if not new_model.strip():
        errors.append("New Phone Model is required.")
    if not validate_imei(new_imei):
        errors.append("New Phone IMEI must be exactly 15 digits.")
    if payment_type == "Exchange":
        if not old_model.strip():
            errors.append("Old Phone Model is required for Exchange.")
        if not validate_imei(old_imei):
            errors.append("Old Phone IMEI must be exactly 15 digits.")
    if not total_price.strip() or not total_price.replace(",", "").replace(" ", "").isdigit():
        errors.append("Enter a valid Total Value (numbers only).")
    if not balance.strip() or not balance.replace(",", "").replace(" ", "").isdigit():
        errors.append("Enter a valid Balance Amount (numbers only).")

    if errors:
        for e in errors:
            st.error(e)
    else:
        inv_num = f"INV-{datetime.now().strftime('%Y%m%d%H%M%S')}"
        data = {
            "invoice_number": inv_num,
            "date": str(bill_date),
            "customer_name": customer_name.strip(),
            "customer_address": customer_address.strip(),
            "customer_contact": customer_contact.strip(),
            "customer_email": customer_email.strip(),
            "new_phone_model": new_model.strip(),
            "new_phone_imei": new_imei.strip(),
            "warranty": warranty,
            "payment_type": payment_type,
            "old_phone_model": old_model.strip() if payment_type == "Exchange" else "",
            "old_phone_imei": old_imei.strip() if payment_type == "Exchange" else "",
            "total_price": total_price.strip(),
            "balance": balance.strip(),
            "other_notes": other_notes.strip(),
        }

        # Save to Google Sheets
        try:
            columns_order = [
                "invoice_number",
                "date",
                "customer_name",
                "customer_address",
                "customer_contact",
                "customer_email",
                "new_phone_model",
                "new_phone_imei",
                "warranty",
                "payment_type",
                "old_phone_model",
                "old_phone_imei",
                "total_price",
                "balance",
                "other_notes"
            ]

            try:
                existing_data = conn.read(ttl=0)
                if existing_data is None or existing_data.empty or (len(existing_data.columns) == 0):
                    existing_data = pd.DataFrame(columns=columns_order)
                else:
                    for col in columns_order:
                        if col not in existing_data.columns:
                            existing_data[col] = ""
                    existing_data = existing_data[columns_order]
            except Exception:
                existing_data = pd.DataFrame(columns=columns_order)

            new_row_df = pd.DataFrame([data], columns=columns_order)
            updated_df = pd.concat([existing_data, new_row_df], ignore_index=True)
            updated_df = updated_df[columns_order]

            conn.update(data=updated_df)
        except Exception as e:
            st.error(f"❌ Failed to save to Google Sheets: {e}")

        # Build PDF & Send
        html = build_pdf_html(data)
        try:
            pdf_bytes = generate_pdf(html)
        except RuntimeError as e:
            st.error(f"❌ PDF generation error: {e}")
            st.stop()

        subject = f"Invoice {inv_num} — {BUSINESS_NAME}"

        with st.spinner("📤 Generating PDF and sending invoice…"):
            ok, err_msg = send_email_with_pdf(data["customer_email"], subject, pdf_bytes, inv_num)

        if ok:
            st.balloons()
            st.success(f"✅ Invoice **{inv_num}** sent as PDF to **{data['customer_email']}**!")
            st.toast("Invoice sent successfully! 🎉", icon="✉️")

            # Offer PDF download in the app as well
            st.download_button(
                label="📥 Download Invoice PDF",
                data=pdf_bytes,
                file_name=f"{inv_num}.pdf",
                mime="application/pdf",
            )
        else:
            st.error(f"❌ Failed to send: {err_msg}")
            st.info("The invoice was still saved to the Google Sheets database.")

# ── SIDEBAR: Sales History ──
with st.sidebar:
    st.markdown("### 📊 Sales Database")
    try:
        df = conn.read(ttl=0)
        if df is not None and not df.empty:
            # Reorder columns to the strict layout if columns are present
            cols = [c for c in [
                "invoice_number", "date", "customer_name", "customer_address", 
                "customer_contact", "customer_email", "new_phone_model", 
                "new_phone_imei", "warranty", "payment_type", "old_phone_model", 
                "old_phone_imei", "total_price", "balance", "other_notes"
            ] if c in df.columns]
            if cols:
                df = df[cols]
            st.metric("Total Invoices", len(df))
            st.dataframe(df.sort_index(ascending=False), use_container_width=True, height=400)
        else:
            st.info("No sales recorded yet. Submit your first bill!")
    except Exception as e:
        st.error(f"Error loading sales database: {e}")
        st.info("Please make sure your Google Sheets connection is properly configured.")
