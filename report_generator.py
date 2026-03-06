"""
Treasury Report AI - PDF Report Generator
==========================================
Generates professional Covenant Compliance Certificate PDFs
suitable for presentation to Saudi bank relationship managers.

Uses ReportLab for PDF generation with professional styling.
"""

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm, inch
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
    HRFlowable, PageBreak, Image
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
from datetime import datetime
import os


# Color palette - professional Saudi corporate style
DARK_NAVY = colors.HexColor('#0D1B2A')
NAVY = colors.HexColor('#1B3A5C')
STEEL_BLUE = colors.HexColor('#2E75B6')
LIGHT_BLUE = colors.HexColor('#D6E4F0')
GOLD = colors.HexColor('#C4A35A')
GREEN = colors.HexColor('#2E7D32')
AMBER = colors.HexColor('#F57F17')
RED = colors.HexColor('#C62828')
WHITE = colors.HexColor('#FFFFFF')
LIGHT_GRAY = colors.HexColor('#F5F5F5')
MEDIUM_GRAY = colors.HexColor('#E0E0E0')
DARK_GRAY = colors.HexColor('#424242')


class NumberedCanvas(canvas.Canvas):
    """Custom canvas for page numbers and headers."""
    def __init__(self, *args, **kwargs):
        canvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_number(num_pages)
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def draw_page_number(self, page_count):
        self.setFont("Helvetica", 8)
        self.setFillColor(DARK_GRAY)
        self.drawRightString(
            A4[0] - 20*mm, 12*mm,
            f"Page {self._pageNumber} of {page_count}"
        )
        self.drawString(
            20*mm, 12*mm,
            "CONFIDENTIAL - Treasury Report AI"
        )
        # Top line
        self.setStrokeColor(GOLD)
        self.setLineWidth(0.5)
        self.line(20*mm, A4[1] - 15*mm, A4[0] - 20*mm, A4[1] - 15*mm)
        # Bottom line
        self.line(20*mm, 18*mm, A4[0] - 20*mm, 18*mm)


def generate_pdf_report(analysis: dict, output_path: str) -> str:
    """
    Generate professional Covenant Compliance Certificate PDF.

    Args:
        analysis: dict from covenant_engine.run_covenant_analysis()
        output_path: path to save the PDF

    Returns:
        Path to generated PDF
    """
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        topMargin=22*mm,
        bottomMargin=25*mm,
        leftMargin=20*mm,
        rightMargin=20*mm,
    )

    styles = getSampleStyleSheet()

    # Custom styles
    styles.add(ParagraphStyle(
        name='ReportTitle',
        fontName='Helvetica-Bold',
        fontSize=22,
        textColor=DARK_NAVY,
        alignment=TA_CENTER,
        spaceAfter=4*mm,
    ))
    styles.add(ParagraphStyle(
        name='ReportSubtitle',
        fontName='Helvetica',
        fontSize=12,
        textColor=STEEL_BLUE,
        alignment=TA_CENTER,
        spaceAfter=2*mm,
    ))
    styles.add(ParagraphStyle(
        name='SectionHeader',
        fontName='Helvetica-Bold',
        fontSize=14,
        textColor=NAVY,
        spaceBefore=8*mm,
        spaceAfter=4*mm,
        borderPadding=(0, 0, 2, 0),
    ))
    styles.add(ParagraphStyle(
        name='SubSection',
        fontName='Helvetica-Bold',
        fontSize=11,
        textColor=STEEL_BLUE,
        spaceBefore=5*mm,
        spaceAfter=3*mm,
    ))
    styles.add(ParagraphStyle(
        name='BodyText2',
        fontName='Helvetica',
        fontSize=9,
        textColor=DARK_GRAY,
        spaceAfter=2*mm,
    ))
    styles.add(ParagraphStyle(
        name='FooterText',
        fontName='Helvetica-Oblique',
        fontSize=8,
        textColor=DARK_GRAY,
        alignment=TA_CENTER,
    ))

    elements = []
    summary = analysis['summary']
    results = analysis['results']
    facilities = analysis['facilities']

    # -- TITLE PAGE / HEADER --
    elements.append(Spacer(1, 10*mm))
    elements.append(Paragraph("COVENANT COMPLIANCE CERTIFICATE", styles['ReportTitle']))
    elements.append(HRFlowable(width="60%", thickness=2, color=GOLD, spaceAfter=5*mm, spaceBefore=2*mm, hAlign='CENTER'))
    elements.append(Paragraph(summary['company_name'], styles['ReportSubtitle']))
    elements.append(Paragraph(f"Reporting Period: {summary['reporting_date']}", styles['ReportSubtitle']))
    elements.append(Paragraph(
        f"Certificate Date: {datetime.now().strftime('%d %B %Y')}",
        styles['ReportSubtitle']
    ))
    elements.append(Spacer(1, 5*mm))

    # Overall status badge
    status = summary['overall_status']
    status_color = GREEN if status == "COMPLIANT" else (AMBER if status == "WARNING" else RED)
    status_table = Table(
        [[Paragraph(
            f"<b>Overall Compliance Status: {status}</b>",
            ParagraphStyle('StatusBadge', fontName='Helvetica-Bold', fontSize=14,
                          textColor=WHITE, alignment=TA_CENTER)
        )]],
        colWidths=[120*mm],
    )
    status_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), status_color),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LEFTPADDING', (0, 0), (-1, -1), 15),
        ('RIGHTPADDING', (0, 0), (-1, -1), 15),
        ('ROUNDEDCORNERS', [4, 4, 4, 4]),
    ]))
    elements.append(status_table)
    elements.append(Spacer(1, 3*mm))

    # Summary counts
    count_data = [
        ['Compliant', 'Warning', 'Breach', 'Total'],
        [str(summary['compliant']), str(summary['warning']),
         str(summary['breach']), str(summary['total_covenants'])],
    ]
    count_table = Table(count_data, colWidths=[35*mm]*4)
    count_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), NAVY),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, MEDIUM_GRAY),
        ('BACKGROUND', (0, 1), (0, 1), colors.HexColor('#E8F5E9')),
        ('BACKGROUND', (1, 1), (1, 1), colors.HexColor('#FFF8E1')),
        ('BACKGROUND', (2, 1), (2, 1), colors.HexColor('#FFEBEE')),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, 1), 16),
    ]))
    elements.append(count_table)
    elements.append(Spacer(1, 8*mm))

    # -- SECTION 1: COVENANT DETAILS BY BANK --
    elements.append(Paragraph("1. FINANCIAL COVENANT COMPLIANCE DETAIL", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=1, color=STEEL_BLUE, spaceAfter=4*mm))

    # Group results by bank
    banks = {}
    for r in results:
        if r.bank not in banks:
            banks[r.bank] = []
        banks[r.bank].append(r)

    for bank_name, bank_results in banks.items():
        elements.append(Paragraph(bank_name, styles['SubSection']))

        header = ['Covenant', 'Required', 'Actual', 'Headroom', 'Status']
        data = [header]

        for r in bank_results:
            status_text = r.status
            headroom_text = f"{r.headroom_pct:+.1f}%"
            data.append([
                r.covenant_name,
                r.threshold_str,
                r.actual_display,
                headroom_text,
                status_text,
            ])

        col_widths = [52*mm, 30*mm, 28*mm, 25*mm, 28*mm]
        table = Table(data, colWidths=col_widths)

        style_commands = [
            ('BACKGROUND', (0, 0), (-1, 0), NAVY),
            ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('FONTSIZE', (0, 1), (-1, -1), 9),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('GRID', (0, 0), (-1, -1), 0.5, MEDIUM_GRAY),
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WHITE, LIGHT_GRAY]),
        ]

        # Color-code status cells
        for i, r in enumerate(bank_results, 1):
            if r.status == "COMPLIANT":
                bg = colors.HexColor('#E8F5E9')
                fg = GREEN
            elif r.status == "WARNING":
                bg = colors.HexColor('#FFF8E1')
                fg = AMBER
            else:
                bg = colors.HexColor('#FFEBEE')
                fg = RED
            style_commands.append(('BACKGROUND', (4, i), (4, i), bg))
            style_commands.append(('TEXTCOLOR', (4, i), (4, i), fg))
            style_commands.append(('FONTNAME', (4, i), (4, i), 'Helvetica-Bold'))

        table.setStyle(TableStyle(style_commands))
        elements.append(table)
        elements.append(Spacer(1, 4*mm))

    # -- SECTION 2: FACILITY UTILIZATION --
    elements.append(Paragraph("2. FACILITY UTILIZATION SUMMARY", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=1, color=STEEL_BLUE, spaceAfter=4*mm))

    fac_header = ['Bank', 'Type', 'Limit\n(SAR M)', 'Outstanding\n(SAR M)', 'Available\n(SAR M)', 'Utilization', 'Maturity']
    fac_data = [fac_header]

    total_limit = 0
    total_outstanding = 0
    total_available = 0

    for f in facilities:
        fac_data.append([
            f.bank.replace('Saudi National Bank (SNB)', 'SNB')
                  .replace('Saudi British Bank (SABB)', 'SABB')
                  .replace('Al Rajhi Bank', 'Al Rajhi'),
            f.facility_type,
            f"{f.limit:.0f}",
            f"{f.outstanding:.0f}",
            f"{f.available:.0f}",
            f"{f.utilization_pct:.1f}%",
            f.maturity_date,
        ])
        total_limit += f.limit
        total_outstanding += f.outstanding
        total_available += f.available

    # Totals row
    total_util = (total_outstanding / total_limit * 100) if total_limit > 0 else 0
    fac_data.append([
        'TOTAL', '', f"{total_limit:.0f}", f"{total_outstanding:.0f}",
        f"{total_available:.0f}", f"{total_util:.1f}%", ''
    ])

    fac_col_widths = [28*mm, 25*mm, 22*mm, 25*mm, 22*mm, 22*mm, 24*mm]
    fac_table = Table(fac_data, colWidths=fac_col_widths)
    fac_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), NAVY),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('FONTNAME', (0, 1), (-1, -2), 'Helvetica'),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('BACKGROUND', (0, -1), (-1, -1), LIGHT_BLUE),
        ('ALIGN', (2, 0), (-1, -1), 'CENTER'),
        ('ALIGN', (0, 0), (1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ('GRID', (0, 0), (-1, -1), 0.5, MEDIUM_GRAY),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [WHITE, LIGHT_GRAY]),
    ]))
    elements.append(fac_table)
    elements.append(Spacer(1, 8*mm))

    # -- SECTION 3: KEY FINANCIAL METRICS --
    fin = analysis['financials']
    elements.append(Paragraph("3. KEY FINANCIAL METRICS", styles['SectionHeader']))
    elements.append(HRFlowable(width="100%", thickness=1, color=STEEL_BLUE, spaceAfter=4*mm))

    metrics_data = [
        ['Metric', 'Value (SAR M)', 'Metric', 'Value (SAR M)'],
        ['Revenue', f"{fin.revenue:.0f}", 'Total Assets', f"{fin.total_assets:.0f}"],
        ['EBITDA', f"{fin.ebitda:.0f}", 'Total Liabilities', f"{fin.total_liabilities:.0f}"],
        ['EBIT', f"{fin.ebit:.0f}", 'Total Equity', f"{fin.total_equity:.0f}"],
        ['Net Income', f"{fin.net_income:.0f}", 'Tangible Net Worth', f"{fin.tangible_net_worth:.0f}"],
        ['Finance Charges', f"{fin.finance_charges:.0f}", 'Current Assets', f"{fin.total_current_assets:.0f}"],
        ['EBITDA Margin', f"{fin.ebitda/fin.revenue*100:.1f}%", 'Current Liabilities', f"{fin.total_current_liabilities:.0f}"],
    ]

    met_table = Table(metrics_data, colWidths=[42*mm, 28*mm, 42*mm, 28*mm])
    met_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), NAVY),
        ('TEXTCOLOR', (0, 0), (-1, 0), WHITE),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('ALIGN', (1, 0), (1, -1), 'RIGHT'),
        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('GRID', (0, 0), (-1, -1), 0.5, MEDIUM_GRAY),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [WHITE, LIGHT_GRAY]),
        ('FONTNAME', (1, 1), (1, -1), 'Helvetica-Bold'),
        ('FONTNAME', (3, 1), (3, -1), 'Helvetica-Bold'),
    ]))
    elements.append(met_table)
    elements.append(Spacer(1, 12*mm))

    # -- CERTIFICATION FOOTER --
    elements.append(HRFlowable(width="100%", thickness=1, color=GOLD, spaceAfter=5*mm))
    elements.append(Paragraph(
        "This Covenant Compliance Certificate has been prepared in accordance with the terms "
        "and conditions of the respective credit facility agreements. The financial data herein "
        "is derived from the unaudited management accounts of the Company for the period indicated above.",
        styles['BodyText2']
    ))
    elements.append(Spacer(1, 8*mm))

    # Signature block
    sig_data = [
        ['Prepared By:', 'Reviewed By:', 'Approved By:'],
        ['', '', ''],
        ['____________________', '____________________', '____________________'],
        ['Treasury Manager', 'VP Finance', 'Chief Financial Officer'],
        [f'Date: {datetime.now().strftime("%d/%m/%Y")}',
         f'Date: {datetime.now().strftime("%d/%m/%Y")}',
         f'Date: {datetime.now().strftime("%d/%m/%Y")}'],
    ]
    sig_table = Table(sig_data, colWidths=[53*mm]*3)
    sig_table.setStyle(TableStyle([
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('TEXTCOLOR', (0, 0), (-1, -1), DARK_GRAY),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 1), (-1, 1), 20),
    ]))
    elements.append(sig_table)
    elements.append(Spacer(1, 8*mm))

    elements.append(Paragraph(
        "This certificate is generated by <b>Treasury Report AI</b> | "
        f"Generated on {datetime.now().strftime('%d %B %Y at %H:%M')}",
        styles['FooterText']
    ))

    # Build PDF
    doc.build(elements, canvasmaker=NumberedCanvas)
    return output_path