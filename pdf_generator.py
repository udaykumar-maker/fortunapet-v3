"""
PDF generator for Proforma Invoices / Quotations — multi-item support.
"""
import io, os
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_LEFT, TA_CENTER

COMPANY_NAME = "Nishant Mouldings Pvt Ltd."
COMPANY_ADDRESS_LINES = [
    "#2, Eralinganna Indl Estate, Srigandakaval,",
    "Vishwaneedam Post, Sunkadakatte, Bangalore-560 091.",
]
COMPANY_URL   = "www.fortunapet.com"
COMPANY_EMAIL = "info@fortunapet.com"

TERMS = [
    "100% Advance payment",
    "Prices - As per the Price increase in the Raw Material.",
]

BANK = {
    "Beneficiary name": "NISHANT MOULDINGS PVT LTD",
    "Bank name":        "SBI BANK",
    "A/c number":       "41107420952",
    "IFSC Code":        "SBIN0008577",
    "Branch":           "Kumara Park",
}

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static", "logo.png")


def generate_pi_pdf(doc):
    buffer = io.BytesIO()
    page = SimpleDocTemplate(buffer, pagesize=A4,
        topMargin=12*mm, bottomMargin=12*mm, leftMargin=15*mm, rightMargin=15*mm)

    styles = getSampleStyleSheet()
    small  = ParagraphStyle('small', parent=styles['Normal'], fontSize=8.5, leading=11, textColor=colors.HexColor('#444444'))
    normal = ParagraphStyle('normal', parent=styles['Normal'], fontSize=9.5, leading=13)
    title_s = ParagraphStyle('title', parent=styles['Heading1'], fontSize=14, alignment=TA_CENTER, spaceAfter=2, textColor=colors.HexColor('#1f2430'))
    label_s = ParagraphStyle('label', parent=styles['Normal'], fontSize=9, textColor=colors.HexColor('#6b7280'))
    value_s = ParagraphStyle('value', parent=styles['Normal'], fontSize=10, textColor=colors.HexColor('#1f2430'))
    bold_s  = ParagraphStyle('bold',  parent=styles['Normal'], fontSize=10, fontName='Helvetica-Bold', textColor=colors.HexColor('#1f2430'))

    elements = []

    # ── HEADER ──
    try:
        logo = Image(LOGO_PATH, width=44*mm, height=44*mm*(108/320)) if os.path.exists(LOGO_PATH) else Paragraph(f"<b>{COMPANY_NAME}</b>", bold_s)
    except Exception:
        logo = Paragraph(f"<b>{COMPANY_NAME}</b>", bold_s)

    company_block = [Paragraph(f"<b>{COMPANY_NAME}</b>", value_s)]
    for line in COMPANY_ADDRESS_LINES:
        company_block.append(Paragraph(line, small))
    company_block.append(Paragraph(f"URL: {COMPANY_URL} | Email: {COMPANY_EMAIL}", small))

    header_tbl = Table([[logo, company_block]], colWidths=[48*mm, 122*mm])
    header_tbl.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP')]))
    elements += [header_tbl, Spacer(1,6), HRFlowable(width="100%",thickness=1,color=colors.HexColor('#d1d5db')), Spacer(1,8)]

    # ── TITLE ──
    doc_label = "PROFORMA INVOICE" if doc.doc_type == "PI" else "QUOTATION"
    elements += [Paragraph(doc_label, title_s), Spacer(1,8)]

    # ── META ──
    meta_left = Table([
        [Paragraph("Quotation No.", label_s), Paragraph(doc.quote_no, value_s)],
        [Paragraph("Date", label_s), Paragraph(doc.doc_date.strftime('%d %b %Y'), value_s)],
        [Paragraph("Dispatch From", label_s), Paragraph(doc.dispatch_from or '—', value_s)],
    ], colWidths=[32*mm, 44*mm])
    meta_left.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP'),('BOTTOMPADDING',(0,0),(-1,-1),4)]))

    bill_to = [Paragraph("<b>Bill To</b>", bold_s), Paragraph(doc.customer.name, value_s)]
    if doc.customer.contact_person: bill_to.append(Paragraph(f"Attn: {doc.customer.contact_person}", small))
    if doc.customer.phone:          bill_to.append(Paragraph(f"Phone: {doc.customer.phone}", small))

    meta_tbl = Table([[meta_left, bill_to]], colWidths=[78*mm, 92*mm])
    meta_tbl.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP')]))
    elements += [meta_tbl, Spacer(1,12)]

    # ── ITEMS TABLE ──
    col_w = [68*mm, 18*mm, 18*mm, 16*mm, 22*mm, 28*mm]
    header_row = ["Description", "Pkg/Box", "Qty", "UOM", "Rate (Rs.)", "Amount (Rs.)"]

    # Get items — use line_items if available, else fall back to legacy single item
    line_items = doc.line_items if doc.line_items else []
    if line_items:
        item_rows = []
        for li in line_items:
            item_rows.append([
                Paragraph(li.item_desc or '', normal),
                li.packaging or '—',
                f"{li.qty:,.0f}" if li.qty else '—',
                li.uom or '—',
                f"{li.price:,.2f}" if li.price else '—',
                f"{li.line_total:,.2f}",
            ])
    else:
        item_rows = [[
            Paragraph(doc.item_desc or '—', normal),
            doc.packaging or '—',
            f"{doc.qty:,.0f}" if doc.qty else '—',
            doc.uom or '—',
            f"{doc.price:,.2f}" if doc.price else '—',
            f"{doc.base_amount:,.2f}",
        ]]

    tbl_data  = [header_row] + item_rows
    item_tbl  = Table(tbl_data, colWidths=col_w)
    item_style = [
        ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1f2430')),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,0),9),
        ('ALIGN',(1,0),(-1,-1),'CENTER'),
        ('ALIGN',(4,0),(-1,-1),'RIGHT'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#d1d5db')),
        ('FONTSIZE',(0,1),(-1,-1),9.5),
        ('TOPPADDING',(0,0),(-1,-1),6),
        ('BOTTOMPADDING',(0,0),(-1,-1),6),
        ('LEFTPADDING',(0,0),(-1,-1),6),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, colors.HexColor('#f8fafc')]),
    ]
    item_tbl.setStyle(TableStyle(item_style))
    elements += [item_tbl, Spacer(1,4)]

    # ── TOTALS ──
    totals = [["Base Amount", f"Rs. {doc.base_amount:,.2f}"]]
    if doc.gst_applied:
        totals.append(["GST @ 18%", f"Rs. {doc.gst_amount:,.2f}"])
    if doc.freight_charges:
        totals.append(["Freight Charges", f"Rs. {doc.freight_charges:,.2f}"])
    totals.append(["Total Payable", f"Rs. {doc.total_amount:,.2f}"])

    tot_tbl = Table(totals, colWidths=[42*mm, 36*mm])
    tot_tbl.setStyle(TableStyle([
        ('ALIGN',(0,0),(-1,-1),'RIGHT'),
        ('FONTSIZE',(0,0),(-1,-1),9.5),
        ('TOPPADDING',(0,0),(-1,-1),4),
        ('BOTTOMPADDING',(0,0),(-1,-1),4),
        ('LINEABOVE',(0,-1),(-1,-1),0.75,colors.HexColor('#1f2430')),
        ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
        ('FONTSIZE',(0,-1),(-1,-1),11),
        ('TOPPADDING',(0,-1),(-1,-1),6),
    ]))
    wrap = Table([["", tot_tbl]], colWidths=[92*mm, 78*mm])
    wrap.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP')]))
    elements += [wrap, Spacer(1,16)]

    if doc.notes:
        elements += [Paragraph("<b>Notes</b>", bold_s), Paragraph(doc.notes, small), Spacer(1,10)]

    # ── TERMS ──
    elements.append(Paragraph("Note: Terms & Conditions", bold_s))
    for i, t in enumerate(TERMS, 1):
        elements.append(Paragraph(f"{i}. {t}", small))
    elements.append(Spacer(1,10))

    # ── BANK ──
    elements.append(Paragraph("Bank Account Details", bold_s))
    bank_rows = [[Paragraph(k, small), Paragraph(v, small)] for k, v in BANK.items()]
    bank_tbl = Table(bank_rows, colWidths=[42*mm, 80*mm])
    bank_tbl.setStyle(TableStyle([('TOPPADDING',(0,0),(-1,-1),1),('BOTTOMPADDING',(0,0),(-1,-1),1)]))
    elements += [bank_tbl, Spacer(1,18)]

    elements.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#e2e6ec')))
    elements.append(Spacer(1,5))
    elements.append(Paragraph("This is a system-generated document.",
        ParagraphStyle('footer', parent=small, alignment=TA_CENTER, textColor=colors.HexColor('#9ca3af'))))

    page.build(elements)
    buffer.seek(0)
    return buffer
