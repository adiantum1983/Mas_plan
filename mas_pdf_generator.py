import io
import os
import pandas as pd
from reportlab.lib.pagesizes import A4, landscape  # type: ignore
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak  # type: ignore
from reportlab.lib import colors  # type: ignore
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle  # type: ignore
from reportlab.pdfbase import pdfmetrics  # type: ignore
from reportlab.pdfbase.ttfonts import TTFont  # type: ignore

def register_fonts():
    font_name = "IPAexGothic"
    try:
        font_path = os.path.join(os.path.dirname(__file__), "ipaexg.ttf")
        if os.path.exists(font_path):
            pdfmetrics.registerFont(TTFont(font_name, font_path))
        else:
            if os.path.exists("C:\\Windows\\Fonts\\meiryo.ttc"):
                pdfmetrics.registerFont(TTFont(font_name, "C:\\Windows\\Fonts\\meiryo.ttc"))
            else:
                font_name = "Helvetica"
    except Exception:
        font_name = "Helvetica"
    return font_name

def format_num(val):
    if pd.isna(val) or val is None: return "0"
    return f"{int(val):,}"

def format_pct(val):
    if pd.isna(val) or val is None: return "0%"
    # Remove the % string if it already contains it, to avoid %%, otherwise format.
    val_str = str(val)
    if "%" in val_str: return val_str
    try:
        return f"{float(val):.1%}"
    except:
        return val_str

def generate_pdf(full_5y_df, df_5y_summary, bs_df, cf_df):
    font_name = register_fonts()
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    
    elements = []
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name="TitleStyle",
        parent=styles["Heading1"],
        fontName=font_name,
        fontSize=18,
        spaceAfter=20,
        alignment=1 # Center
    )
    h2_style = ParagraphStyle(
        name="H2Style",
        parent=styles["Heading2"],
        fontName=font_name,
        fontSize=14,
        spaceAfter=10,
        spaceBefore=15
    )
    
    elements.append(Paragraph("継続MAS 5カ年計画レポート", title_style))
    
    # 1. 5Y Summary
    elements.append(Paragraph("1. 5カ年計画 総合指標", h2_style))
    summary_data = [["指標"] + [f"第{i}期" for i in range(1, 6)]]
    for row_name in df_5y_summary.columns:
        row_vals = [row_name]
        for val in df_5y_summary[row_name].values:
            if row_name == '変動費率':
                row_vals.append(format_pct(val))
            else:
                row_vals.append(format_num(val))
        summary_data.append(row_vals)
        
    t_summary = Table(summary_data, colWidths=[120] + [100]*5)
    t_summary.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#34495e")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 6),
    ]))
    elements.append(t_summary)
    
    # 2. Granular Details
    elements.append(PageBreak())
    elements.append(Paragraph("2. 5カ年計画 明細科目", h2_style))
    
    if not full_5y_df.empty:
        # Group by 'Cat' (管理科目) while preserving sequence structure
        cats_in_df = full_5y_df['Cat'].unique()
        
        for cat in cats_in_df:
            cat_df = full_5y_df[full_5y_df['Cat'] == cat]
            
            elements.append(Paragraph(f"■ {cat}", ParagraphStyle(name="CatTitle", fontName=font_name, fontSize=11, spaceBefore=10, spaceAfter=5)))
            
            detail_data = [["勘定科目名", "第1期", "第2期", "第3期", "第4期", "第5期"]]
            for _, row in cat_df.iterrows():
                row_vals = [str(row['勘定科目名'])]
                for i in range(1, 6):
                    row_vals.append(format_num(row[f'第{i}期']))
                detail_data.append(row_vals)
                
            # Add Subtotal Row
            subtotal_vals = [f"【{cat} 合計】"]
            for i in range(1, 6):
                subtotal_vals.append(format_num(cat_df[f'第{i}期'].sum()))
            detail_data.append(subtotal_vals)
                
            t_detail = Table(detail_data, colWidths=[200] + [90]*5)
            t_detail.setStyle(TableStyle([
                ('FONTNAME', (0,0), (-1,-1), font_name),
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#ecf0f1")),
                ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#d5f5e3")), # Subtotal row color
                ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
                ('ALIGN', (0,0), (-1,0), 'CENTER'),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('FONTSIZE', (0,0), (-1,-1), 9),
            ]))
            elements.append(t_detail)
            
    # 3. B/S and CF
    elements.append(PageBreak())
    elements.append(Paragraph("3. 予測貸借対照表 (B/S) ピックアップ", h2_style))
    
    bs_data = [["項目"] + [f"第{i}期" for i in range(1, 6)]]
    for row_name in bs_df.columns:
        row_vals = [row_name]
        for val in bs_df[row_name].values:
            row_vals.append(format_num(val))
        bs_data.append(row_vals)
        
    t_bs = Table(bs_data, colWidths=[120] + [100]*5)
    t_bs.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#34495e")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
    ]))
    elements.append(t_bs)
    
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("4. キャッシュ・フロー予測", h2_style))
    
    cf_data = [["項目"] + [f"第{i}期" for i in range(1, 6)]]
    for row_name in cf_df.columns:
        row_vals = [row_name]
        for val in cf_df[row_name].values:
            row_vals.append(format_num(val))
        cf_data.append(row_vals)
        
    t_cf = Table(cf_data, colWidths=[150] + [100]*5)
    t_cf.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#34495e")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('BACKGROUND', (0,-1), (-1,-1), colors.HexColor("#fadbd8")), # FCF row
        ('ALIGN', (1,0), (-1,-1), 'RIGHT'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
    ]))
    elements.append(t_cf)
    
    doc.build(elements)
    
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes
