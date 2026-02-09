from pathlib import Path
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.lib.pagesizes import A4
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer

def generate_pdf_report(file_path, scraping_stats, crm_result, year, local_file_path):
    try:
        doc = SimpleDocTemplate(
            file_path,
            pagesize=A4,
            rightMargin=72,
            leftMargin=72,
            topMargin=72,
            bottomMargin=72
        )        
        elements = []        
        styles = getSampleStyleSheet()        
        title_style = ParagraphStyle(
            'CustomTitle',
            parent=styles['Heading1'],
            fontSize=24,
            alignment=TA_CENTER,
            textColor=colors.HexColor('#dc2626'),
            spaceAfter=20
        )        
        subtitle_style = ParagraphStyle(
            'CustomSubtitle',
            parent=styles['Heading2'],
            fontSize=14,
            alignment=TA_CENTER,
            textColor=colors.HexColor('#666666'),
            spaceAfter=30
        )        
        heading_style = ParagraphStyle(
            'CustomHeading',
            parent=styles['Heading2'],
            fontSize=16,
            alignment=TA_LEFT,
            textColor=colors.HexColor('#1e3a8a'),
            spaceBefore=20,
            spaceAfter=10
        )        
        normal_style = ParagraphStyle(
            'CustomNormal',
            parent=styles['Normal'],
            fontSize=11,
            alignment=TA_LEFT,
            textColor=colors.HexColor('#374151'),
            spaceAfter=6
        )        
        list_style = ParagraphStyle(
            'CustomList',
            parent=styles['Normal'],
            fontSize=10,
            alignment=TA_LEFT,
            textColor=colors.HexColor('#4b5563'),
            leftIndent=20,
            spaceAfter=3
        )        
        elements.append(Paragraph("CMDA SCRAPING REPORT", title_style))
        elements.append(Paragraph(f"Ajantha Bathroom Products - {year}", subtitle_style))        
        elements.append(Paragraph(f"<b>Report Generated:</b> {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", normal_style))
        elements.append(Paragraph(f"<b>Data File:</b> {Path(local_file_path).name if local_file_path else 'Not Available'}", normal_style))
        elements.append(Spacer(1, 20))        
        elements.append(Paragraph("1. Scraping Statistics", heading_style))        
        scraping_data = [
            ["Metric", "Count"],
            ["Total Records Attempted", str(scraping_stats.get('total_attempted', 0))],
            ["Successfully Scraped", str(scraping_stats.get('successful_scraped', 0))],
            ["Failed to Scrape", str(scraping_stats.get('failed_scraped', 0))],
        ]        
        scraping_table = Table(scraping_data, colWidths=[3*inch, 2*inch])
        scraping_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#dc2626')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8f9fa')),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 11),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))        
        elements.append(scraping_table)
        elements.append(Spacer(1, 20))        
        failed_file_numbers = scraping_stats.get('failed_file_numbers', [])
        if failed_file_numbers:
            elements.append(Paragraph("<b>Failed to Scrape File Numbers:</b>", normal_style))
            for file_no in failed_file_numbers[:20]:
                elements.append(Paragraph(f"• {file_no}", list_style))
            if len(failed_file_numbers) > 20:
                elements.append(Paragraph(f"... and {len(failed_file_numbers) - 20} more", list_style))
        else:
            elements.append(Paragraph("<b>Failed to Scrape File Numbers:</b> None", normal_style))        
        elements.append(Spacer(1, 30))        
        elements.append(Paragraph("2. CRM Integration Results", heading_style))        
        crm_status = crm_result.get('status', False)
        crm_message = crm_result.get('message', 'No message')
        analysis_data = crm_result.get('analysis_data', {})        
        crm_data = [
            ["CRM Import Status", "Successful" if crm_status else "Failed"],
            ["Message", crm_message],
            ["Status Code", str(crm_result.get('statusCode', 'N/A'))]
        ]        
        crm_table = Table(crm_data, colWidths=[2*inch, 4*inch])
        crm_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1e3a8a')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f8f9fa')),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 11),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))        
        elements.append(crm_table)
        elements.append(Spacer(1, 30))        
        elements.append(Paragraph("3. Data Analysis", heading_style))        
        new_records_count = analysis_data.get('new_records_count', 0)
        elements.append(Paragraph(f"<b>New Records Count:</b> {new_records_count}", normal_style))
        new_file_numbers = analysis_data.get('new_file_numbers', [])
        if new_file_numbers:
            elements.append(Paragraph("<b>New Records File Numbers:</b>", normal_style))
            for file_no in new_file_numbers[:20]:
                elements.append(Paragraph(f"• {file_no}", list_style))
            if len(new_file_numbers) > 20:
                elements.append(Paragraph(f"... and {len(new_file_numbers) - 20} more", list_style))
        elements.append(Spacer(1, 20))        
        matched_count = analysis_data.get('matched_count', 0)
        unmatched_count = analysis_data.get('unmatched_count', 0)        
        match_data = [
            ["Record Type", "Count"],
            ["Matched Records", str(matched_count)],
            ["Unmatched Records", str(unmatched_count)]
        ]        
        match_table = Table(match_data, colWidths=[3*inch, 2*inch])
        match_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#059669')),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor('#f0fdf4')),
            ('GRID', (0, 0), (-1, -1), 1, colors.black),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 11),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))        
        elements.append(match_table)
        elements.append(Spacer(1, 20))        
        matched_file_numbers = analysis_data.get('matched_file_numbers', [])
        if matched_file_numbers:
            elements.append(Paragraph("<b>Matched Record File Numbers:</b>", normal_style))
            for file_no in matched_file_numbers[:20]:
                elements.append(Paragraph(f"• {file_no}", list_style))
            if len(matched_file_numbers) > 20:
                elements.append(Paragraph(f"... and {len(matched_file_numbers) - 20} more", list_style))
        unmatched_file_numbers = analysis_data.get('unmatched_file_numbers', [])
        if unmatched_file_numbers:
            elements.append(Paragraph("<b>Unmatched Record File Numbers:</b>", normal_style))
            for file_no in unmatched_file_numbers[:20]:
                elements.append(Paragraph(f"• {file_no}", list_style))
            if len(unmatched_file_numbers) > 20:
                elements.append(Paragraph(f"... and {len(unmatched_file_numbers) - 20} more", list_style))
        elements.append(Spacer(1, 20))        
        unmatched_areas = analysis_data.get('unmatched_areas', [])
        if unmatched_areas:
            elements.append(Paragraph(f"<b>Unmatched Areas Count:</b> {len(unmatched_areas)}", normal_style))
            elements.append(Paragraph("<b>Unmatched Areas Details:</b>", normal_style))
            for area in unmatched_areas[:10]:
                elements.append(Paragraph(f"• {area}", list_style))
            if len(unmatched_areas) > 10:
                elements.append(Paragraph(f"... and {len(unmatched_areas) - 10} more", list_style))
        else:
            elements.append(Paragraph("<b>Unmatched Areas Count:</b> 0", normal_style))
        elements.append(Spacer(1, 40))        
        footer_style = ParagraphStyle(
            'Footer',
            parent=styles['Normal'],
            fontSize=9,
            alignment=TA_CENTER,
            textColor=colors.HexColor('#666666')
        )        
        elements.append(Paragraph("Generated by Ajantha Bathroom Products Automation System", footer_style))
        elements.append(Paragraph(f"Report ID: {datetime.now().strftime('%Y%m%d%H%M%S')}", footer_style))
        doc.build(elements)
        return True        
    except Exception as e:
        print(f"❌ Error generating PDF report: {e}")
        return False