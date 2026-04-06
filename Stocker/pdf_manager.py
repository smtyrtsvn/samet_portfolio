from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from datetime import datetime


class PDFGenerator:
    def __init__(self):
        try:
            # Türkçe karakter desteği için Arial fontunu kullanmayı dener
            # Bilgisayarda 'arial.ttf' bulunmazsa hata verebilir, o yüzden try-except var.
            pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
            self.font_name = 'Arial'
        except:
            self.font_name = 'Helvetica'

    def create_invoice(self, filename, invoice_no, type_text, items):
        doc = SimpleDocTemplate(filename, pagesize=A4)
        elements = []
        styles = getSampleStyleSheet()

        title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontName=self.font_name, alignment=1)
        normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontName=self.font_name)

        elements.append(Paragraph(f"FATURA / {type_text}", title_style))
        elements.append(Spacer(1, 20))

        elements.append(Paragraph(f"<b>Fatura No:</b> {invoice_no}", normal_style))
        elements.append(Paragraph(f"<b>Tarih:</b> {datetime.now().strftime('%d-%m-%Y')}", normal_style))
        elements.append(Spacer(1, 20))

        # Tablo Başlıkları
        data = [['SKU', 'Model', 'Varyant', 'Adet', 'Birim Fiyat', 'Toplam']]

        total_sum = 0
        for item in items:
            row = [
                item['sku'],
                item['model'],
                item['var'],
                str(item['qty']),
                f"{item['price']:.2f}",
                f"{item['total']:.2f}"
            ]
            data.append(row)
            total_sum += item['total']

        data.append(['', '', '', '', 'GENEL TOPLAM:', f"{total_sum:.2f}"])

        # Tablo Stili
        table = Table(data, colWidths=[100, 120, 100, 50, 70, 80])
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), self.font_name),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (-2, -1), (-1, -1), self.font_name),
            ('FONTSIZE', (-2, -1), (-1, -1), 12),
            ('TEXTCOLOR', (-2, -1), (-1, -1), colors.darkblue),
        ]))

        elements.append(table)
        elements.append(Spacer(1, 30))
        elements.append(Paragraph("Bu belge bilgisayar ortamında oluşturulmuştur.", normal_style))

        doc.build(elements)