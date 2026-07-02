"""
PDF generator for Proforma Invoices / Quotations.
Builds a branded PDF on the fly using reportlab.
"""
import io
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, HRFlowable
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_LEFT, TA_CENTER

COMPANY_NAME = "Nishant Mouldings Pvt Ltd."
COMPANY_ADDRESS_LINES = [
    "#2, Eralinganna Indl Estate, Srigandakaval,",
    "Vishwaneedam Post, Sunkadakatte, Bangalore-560 091.",
]
COMPANY_URL = "www.fortunapet.com"
COMPANY_EMAIL = "info@fortunapet.com"

TERMS_AND_CONDITIONS = [
    "100% Advance payment",
    "Prices - As per the Price increase in the Raw Material.",
]

BANK_DETAILS = {
    "Beneficiary name": "NISHANT MOULDINGS PVT LTD",
    "Bank name": "SBI BANK",
    "A/c number": "41107420952",
    "IFSC Code": "SBIN0008577",
    "Branch": "Kumara Park",
}

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "logo.png")


def generate_pi_pdf(doc):
    """
    doc: a Document model instance (with .customer relationship loaded)
    Returns: BytesIO buffer containing the PDF
    """
    buffer = io.BytesIO()
    page = SimpleDocTemplate(
        buffer, pagesize=A4,
        topMargin=14 * mm, bottomMargin=14 * mm,
        leftMargin=16 * mm, rightMargin=16 * mm,
    )

    styles = getSampleStyleSheet()
    small = ParagraphStyle('small', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor('#444444'))
    normal = ParagraphStyle('normal', parent=styles['Normal'], fontSize=9.5, leading=13)
    title_style = ParagraphStyle('docTitle', parent=styles['Heading1'], fontSize=15, alignment=TA_CENTER, spaceAfter=2, textColor=colors.HexColor('#1f2430'))
    label_style = ParagraphStyle('label', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#6b7280'))
    value_style = ParagraphStyle('value', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#1f2430'))
    heading_style = ParagraphStyle('heading', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#1f2430'), fontName='Helvetica-Bold')

    elements = []

    # ---- HEADER: logo + company info ----
    try:
        if os.path.exists(LOGO_PATH):
            logo = Image(LOGO_PATH, width=46 * mm, height=46 * mm * (108 / 320))
        else:
            logo = Paragraph(f"<b>{COMPANY_NAME}</b>", heading_style)
    except Exception:
        logo = Paragraph(f"<b>{COMPANY_NAME}</b>", heading_style)

    company_block = [
        Paragraph(f"<b>{COMPANY_NAME}</b>", value_style),
    ]
    for line in COMPANY_ADDRESS_LINES:
        company_block.append(Paragraph(line, small))
    company_block.append(Paragraph(f"URL: {COMPANY_URL} | Email: {COMPANY_EMAIL}", small))

    header_table = Table(
        [[logo, company_block]],
        colWidths=[50 * mm, 124 * mm],
    )
    header_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
    ]))
    elements.append(header_table)
    elements.append(Spacer(1, 8))
    elements.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#d1d5db')))
    elements.append(Spacer(1, 10))

    # ---- DOCUMENT TITLE ----
    doc_label = "PROFORMA INVOICE" if doc.doc_type == "PI" else "QUOTATION"
    elements.append(Paragraph(doc_label, title_style))
    elements.append(Spacer(1, 8))

    # ---- META INFO: doc no / date / dispatch  +  bill-to ----
    meta_left = [
        [Paragraph("Quotation No.", label_style), Paragraph(doc.quote_no, value_style)],
        [Paragraph("Date", label_style), Paragraph(doc.doc_date.strftime('%d %b %Y'), value_style)],
        [Paragraph("Dispatch From", label_style), Paragraph(doc.dispatch_from or '—', value_style)],
    ]
    meta_left_table = Table(meta_left, colWidths=[32 * mm, 45 * mm])
    meta_left_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))

    bill_to_lines = [Paragraph("<b>Bill To</b>", heading_style)]
    bill_to_lines.append(Paragraph(doc.customer.name, value_style))
    if doc.customer.contact_person:
        bill_to_lines.append(Paragraph(f"Attn: {doc.customer.contact_person}", small))
    if doc.customer.phone:
        bill_to_lines.append(Paragraph(f"Phone: {doc.customer.phone}", small))
    if doc.customer.email:
        bill_to_lines.append(Paragraph(f"Email: {doc.customer.email}", small))

    meta_table = Table(
        [[meta_left_table, bill_to_lines]],
        colWidths=[80 * mm, 94 * mm],
    )
    meta_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    elements.append(meta_table)
    elements.append(Spacer(1, 14))

    # ---- ITEM TABLE ----
    item_header = ["Description", "Pkg/Box", "Qty", "UOM", "Rate (Rs.)", "Amount (Rs.)"]
    item_row = [
        Paragraph(doc.item_desc, normal),
        doc.packaging or '—',
        f"{doc.qty:,.0f}" if doc.qty else '—',
        doc.uom or '—',
        f"{doc.price:,.2f}" if doc.price else '—',
        f"{doc.base_amount:,.2f}",
    ]
    item_table = Table(
        [item_header, item_row],
        colWidths=[64 * mm, 20 * mm, 20 * mm, 18 * mm, 24 * mm, 28 * mm],
    )
    item_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f2430')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (4, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#d1d5db')),
        ('FONTSIZE', (0, 1), (-1, -1), 9.5),
        ('TOPPADDING', (0, 0), (-1, -1), 7),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 7),
        ('LEFTPADDING', (0, 0), (-1, -1), 6),
    ]))
    elements.append(item_table)
    elements.append(Spacer(1, 4))

    # ---- TOTALS BOX ----
    totals_rows = [["Base Amount", f"Rs. {doc.base_amount:,.2f}"]]
    if doc.gst_applied:
        totals_rows.append(["GST @ 18%", f"Rs. {doc.gst_amount:,.2f}"])
    if doc.freight_charges:
        totals_rows.append(["Freight Charges", f"Rs. {doc.freight_charges:,.2f}"])
    totals_rows.append(["Total Payable", f"Rs. {doc.total_amount:,.2f}"])

    totals_table = Table(totals_rows, colWidths=[40 * mm, 35 * mm])
    style_cmds = [
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('FONTSIZE', (0, 0), (-1, -1), 9.5),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LINEABOVE', (0, -1), (-1, -1), 0.75, colors.HexColor('#1f2430')),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, -1), (-1, -1), 11),
        ('TOPPADDING', (0, -1), (-1, -1), 7),
    ]
    totals_table.setStyle(TableStyle(style_cmds))

    # right-align the totals box on the page
    wrap_table = Table([["", totals_table]], colWidths=[94 * mm, 80 * mm])
    wrap_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP')]))
    elements.append(wrap_table)
    elements.append(Spacer(1, 18))

    if doc.notes:
        elements.append(Paragraph("<b>Notes</b>", heading_style))
        elements.append(Paragraph(doc.notes, small))
        elements.append(Spacer(1, 12))

    # ---- TERMS & CONDITIONS ----
    elements.append(Paragraph("Note: Terms & Conditions", heading_style))
    for i, term in enumerate(TERMS_AND_CONDITIONS, 1):
        elements.append(Paragraph(f"{i}. {term}", small))
    elements.append(Spacer(1, 12))

    # ---- BANK DETAILS ----
    elements.append(Paragraph("Bank Account Details", heading_style))
    bank_rows = [[Paragraph(f"{k}", small), Paragraph(v, small)] for k, v in BANK_DETAILS.items()]
    bank_table = Table(bank_rows, colWidths=[40 * mm, 80 * mm])
    bank_table.setStyle(TableStyle([
        ('TOPPADDING', (0, 0), (-1, -1), 1),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 1),
    ]))
    elements.append(bank_table)
    elements.append(Spacer(1, 20))

    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#e2e6ec')))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(
        "This is a system-generated document and does not require a physical signature.",
        ParagraphStyle('footer', parent=small, alignment=TA_CENTER, textColor=colors.HexColor('#9ca3af'))
    ))

    page.build(elements)
    buffer.seek(0)
    return buffer
