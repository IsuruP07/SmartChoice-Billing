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
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

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




def make_paragraph(text, size=9, bold=False, color=colors.HexColor('#222222'), align=TA_LEFT):
    style_name = f"custom_{size}_{bold}_{color.hexval()}_{align}"
    styles = getSampleStyleSheet()
    if style_name not in styles:
        p_style = ParagraphStyle(
            name=style_name,
            parent=styles['Normal'],
            fontName='Helvetica-Bold' if bold else 'Helvetica',
            fontSize=size,
            leading=size + 3,
            textColor=color,
            alignment=align
        )
        styles.add(p_style)
    else:
        p_style = styles[style_name]
    return Paragraph(text, p_style)


def generate_pdf(d: dict) -> bytes:
    """Convert transaction data dictionary to PDF bytes using reportlab."""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=24,
        rightMargin=24,
        topMargin=24,
        bottomMargin=24
    )
    story = []

    # 1. Header Box
    header_title = make_paragraph("SMART CHOICE MOBILES", size=22, bold=True, color=colors.white, align=TA_CENTER)
    header_sub = make_paragraph("Maligathanna, Gurudeniya    |    Phone / WhatsApp: 0723175373    |    smartchoicemobileskandy@gmail.com", size=9.5, bold=False, color=colors.HexColor('#CCCCCC'), align=TA_CENTER)
    header_table = Table([[header_title], [header_sub]], colWidths=[547])
    header_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#222222')),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,0), 2),
        ('TOPPADDING', (0,0), (-1,-1), 10),
        ('BOTTOMPADDING', (0,1), (-1,-1), 10),
        ('LINEBELOW', (0,1), (-1,1), 5, colors.HexColor('#990000')),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 25))

    # 2. Billed To / Invoice Details Section
    billed_to_flow = [
        make_paragraph("BILLED TO", size=9, bold=True, color=colors.HexColor('#990000')),
        Spacer(1, 4),
        make_paragraph(d['customer_name'], size=11, bold=True),
        Spacer(1, 2),
        make_paragraph(d.get('customer_address', ''), size=9, color=colors.HexColor('#444444')),
        Spacer(1, 1),
        make_paragraph(d.get('customer_contact', ''), size=9, color=colors.HexColor('#444444')),
        Spacer(1, 1),
        make_paragraph(d['customer_email'], size=9, color=colors.HexColor('#444444')),
    ]

    details_data = [
        [make_paragraph("Invoice #:", size=9, bold=True, align=TA_RIGHT), make_paragraph(d['invoice_number'], size=9, color=colors.HexColor('#444444'), align=TA_RIGHT)],
        [make_paragraph("Date:", size=9, bold=True, align=TA_RIGHT), make_paragraph(d['date'], size=9, color=colors.HexColor('#444444'), align=TA_RIGHT)],
        [make_paragraph("Payment:", size=9, bold=True, align=TA_RIGHT), make_paragraph(d['payment_type'], size=9, color=colors.HexColor('#444444'), align=TA_RIGHT)],
    ]
    details_table = Table(details_data, colWidths=[140, 100])
    details_table.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'RIGHT'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
    ]))

    details_flow = [
        make_paragraph("INVOICE DETAILS", size=9, bold=True, color=colors.HexColor('#990000'), align=TA_RIGHT),
        Spacer(1, 4),
        details_table
    ]

    info_table = Table([[billed_to_flow, details_flow]], colWidths=[273.5, 273.5])
    info_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 20))

    # 3. Transaction Table
    tx_rows = [
        [make_paragraph("DESCRIPTION", size=9, bold=True, color=colors.white),
         make_paragraph("DETAILS", size=9, bold=True, color=colors.white)],
        [make_paragraph("New Phone Model", size=9, bold=True, color=colors.HexColor('#444444')),
         make_paragraph(d['new_phone_model'], size=9, bold=True)],
        [make_paragraph("New Phone IMEI", size=9, bold=True, color=colors.HexColor('#444444')),
         make_paragraph(d['new_phone_imei'], size=9, color=colors.HexColor('#222222'))],
        [make_paragraph("Warranty Period", size=9, bold=True, color=colors.HexColor('#444444')),
         make_paragraph(d['warranty'], size=9)],
        [make_paragraph("Payment Type", size=9, bold=True, color=colors.HexColor('#444444')),
         make_paragraph(d['payment_type'], size=9)],
    ]
    if d['payment_type'] == "Exchange":
        tx_rows.append([
            make_paragraph("Old Phone Model", size=9, bold=True, color=colors.HexColor('#444444')),
            make_paragraph(d['old_phone_model'], size=9)
        ])
        tx_rows.append([
            make_paragraph("Old Phone IMEI", size=9, bold=True, color=colors.HexColor('#444444')),
            make_paragraph(d['old_phone_imei'], size=9)
        ])

    tx_style = [
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#990000')),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#EEEEEE')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#EEEEEE')),
    ]
    for idx in range(1, len(tx_rows)):
        bg_color = colors.HexColor('#FFFFFF') if idx % 2 == 1 else colors.HexColor('#F9F9F9')
        tx_style.append(('BACKGROUND', (0, idx), (-1, idx), bg_color))

    tx_table = Table(tx_rows, colWidths=[200, 347])
    tx_table.setStyle(TableStyle(tx_style))
    story.append(tx_table)
    story.append(Spacer(1, 10))

    # 4. Totals Section
    totals_data = [
        [make_paragraph("Total Amount", size=9, bold=True, color=colors.HexColor('#666666')),
         make_paragraph(f"Rs. {d['total_price']}", size=11, bold=True, align=TA_RIGHT)],
        [make_paragraph("Balance Due", size=9, bold=True, color=colors.HexColor('#222222')),
         make_paragraph(f"Rs. {d['balance']}", size=13, bold=True, color=colors.HexColor('#990000'), align=TA_RIGHT)],
    ]
    totals_table = Table(totals_data, colWidths=[120, 127])
    totals_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ('LINEBELOW', (0,0), (-1,0), 0.5, colors.HexColor('#EEEEEE')),
    ]))

    totals_layout = Table([["", totals_table]], colWidths=[300, 247])
    totals_layout.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(totals_layout)
    story.append(Spacer(1, 10))

    # 5. Notes / Remarks Box (Conditional)
    if d.get("other_notes", "").strip():
        notes_data = [
            [make_paragraph("Notes / Remarks", size=9, bold=True, color=colors.HexColor('#222222'))],
            [make_paragraph(d['other_notes'], size=9, color=colors.HexColor('#444444'))]
        ]
        notes_table = Table(notes_data, colWidths=[547])
        notes_table.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.white),
            ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#EEEEEE')),
            ('TOPPADDING', (0,0), (-1,0), 6),
            ('BOTTOMPADDING', (0,0), (-1,0), 2),
            ('TOPPADDING', (0,1), (-1,-1), 2),
            ('BOTTOMPADDING', (0,1), (-1,-1), 6),
            ('LEFTPADDING', (0,0), (-1,-1), 10),
            ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ]))
        story.append(notes_table)
        story.append(Spacer(1, 10))

    # 6. Terms & Warranty Policy
    terms_body = (
        "1. Warranty claims are valid ONLY for issues that occur strictly within the specified warranty period mentioned in this invoice.<br/>"
        "2. The warranty seal/sticker applied on the device MUST be intact. Any tampering, tearing, or removal of the seal will immediately void the entire warranty.<br/>"
        "3. The original invoice must be presented for any warranty claims.<br/>"
        "4. Physical damage, liquid damage, and display drops are completely not covered under warranty.<br/>"
        "5. Software issues are subject to the standard company policy."
    )
    terms_data = [
        [make_paragraph("Terms & Warranty Policy", size=9, bold=True, color=colors.HexColor('#222222'))],
        [make_paragraph(terms_body, size=8, color=colors.HexColor('#444444'))]
    ]
    terms_table = Table(terms_data, colWidths=[547])
    terms_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F9F9F9')),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#EEEEEE')),
        ('TOPPADDING', (0,0), (-1,0), 5),
        ('BOTTOMPADDING', (0,0), (-1,0), 2),
        ('TOPPADDING', (0,1), (-1,-1), 2),
        ('BOTTOMPADDING', (0,1), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
    ]))
    story.append(terms_table)
    story.append(Spacer(1, 10))

    # 7. Footer
    footer_line = Table([[""]], colWidths=[547])
    footer_line.setStyle(TableStyle([
        ('LINEABOVE', (0,0), (-1,-1), 1.5, colors.HexColor('#990000')),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('TOPPADDING', (0,0), (-1,-1), 0),
    ]))
    story.append(footer_line)
    story.append(Spacer(1, 4))

    footer_text_1 = make_paragraph("Thank you for choosing <b>Smart Choice Mobiles Kandy</b>!", size=8.5, color=colors.HexColor('#555555'), align=TA_CENTER)
    footer_text_2 = make_paragraph("Computer-generated document. No signature required.", size=8, color=colors.HexColor('#999999'), align=TA_CENTER)
    story.append(footer_text_1)
    story.append(Spacer(1, 1))
    story.append(footer_text_2)

    doc.build(story)
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
        try:
            pdf_bytes = generate_pdf(data)
        except Exception as e:
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
