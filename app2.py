import os
import requests
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==========================================
# 0. 日本語フォント登録（環境に合わせて調整）
# ==========================================
FONT_PATH = "NotoSansJP-Regular.ttf"  # 必要に応じてパスを指定
if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont('Japanese', FONT_PATH))
    DEFAULT_FONT = 'Japanese'
else:
    DEFAULT_FONT = 'Helvetica'

def fetch_record_by_order(base_url, api_key, app_id, order_no):
    """@pocketから受注番号（field-2）で1件のレコードを取得する"""
    url = f"{base_url}/api/apps/{app_id}/records"
    headers = {"X-At-Pocket-API-Key": api_key, "Accept": "application/json"}
    params = {"query": f'field-2 = "{order_no}"'}
    
    try:
        response = requests.get(url, headers=headers, params=params, timeout=15)
        if response.status_code == 200:
            result = response.json()
            records = result.get("records", result.get("Data", result.get("Records", [])))
            for item in records:
                rec_body = item.get("record", {})
                merged = {**item, **rec_body}
                if str(merged.get("field-2", "")) == str(order_no):
                    return merged
    except Exception as e:
        print(f"API通信エラー: {e}")
    return None

def generate_seihon_card(data, output_filename="seihon_card.pdf"):
    """製本進行カードのPDFを生成する"""
    doc = SimpleDocTemplate(
        output_filename,
        pagesize=A4,
        rightMargin=10 * mm, leftMargin=10 * mm,
        topMargin=10 * mm, bottomMargin=10 * mm
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CardTitle',
        fontName=DEFAULT_FONT,
        fontSize=18,
        leading=22,
        alignment=1, # 中央揃え
        fontWeight='bold'
    )
    
    normal_style = ParagraphStyle(
        'CardNormal',
        fontName=DEFAULT_FONT,
        fontSize=9,
        leading=12
    )
    
    bold_style = ParagraphStyle(
        'CardBold',
        fontName=DEFAULT_FONT,
        fontSize=9,
        leading=12,
        fontWeight='bold'
    )

    story = []

    # 1. ヘッダー部分（納期・見本・配本・タイトル）
    header_data = [
        [
            Paragraph(f"<b>納期:</b> {data.get('field-52', '')}", normal_style),
            Paragraph(f"<b>見本:</b> {data.get('field-54', '')}", normal_style),
            Paragraph(f"<b>配本:</b> {data.get('field-56', '')}", normal_style),
            Paragraph("製本進行カード", title_style),
            Paragraph(f"<b>受注番号:</b> {data.get('field-2', '')}", bold_style)
        ]
    ]
    header_table = Table(header_data, colWidths=[40*mm, 30*mm, 30*mm, 60*mm, 30*mm])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (3,0), (3,0), colors.lightgrey),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 5 * mm))

    # 2. 基本情報（書名、出版社、印刷所など）
    info_data = [
        [Paragraph("書名", bold_style), Paragraph(data.get('field-13', ''), normal_style),
         Paragraph("出版社", bold_style), Paragraph(data.get('field-8', ''), normal_style)],
        [Paragraph("ISBN", bold_style), Paragraph(data.get('field-93', ''), normal_style),
         Paragraph("印刷所", bold_style), Paragraph(data.get('field-14', ''), normal_style)],
        [Paragraph("種別/刷数", bold_style), Paragraph(f"{data.get('field-19', '')} / {data.get('field-28', '')}刷 (回:{data.get('field-29', '')})", normal_style),
         Paragraph("定価", bold_style), Paragraph(f"¥{data.get('field-20', '')}", normal_style)],
        [Paragraph("総頁数", bold_style), Paragraph(f"{data.get('field-82', '')} 頁", normal_style),
         Paragraph("印刷/製本数", bold_style), Paragraph(f"印:{data.get('field-3', '')}部 / 製:{data.get('field-4', '')}部", normal_style)],
    ]
    info_table = Table(info_data, colWidths=[25*mm, 70*mm, 25*mm, 70*mm])
    info_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
        ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 3 * mm))

    # 3. 仕様情報（仕上仕様、背固仕様、寸法など）
    spec_data = [
        [Paragraph("仕上仕様", bold_style), Paragraph(f"{data.get('field-41', '')} / {data.get('field-42', '')}", normal_style),
         Paragraph("背固仕様", bold_style), Paragraph(data.get('field-43', ''), normal_style)],
        [Paragraph("仕上寸法", bold_style), Paragraph(f"天地:{data.get('field-24', '')} / 左右:{data.get('field-25', '')} ({data.get('field-44', '')})", normal_style),
         Paragraph("付物/見返し", bold_style), Paragraph(f"付:{data.get('field-49', '')} / 見:{data.get('field-48', '')}", normal_style)],
        [Paragraph("口絵 / 折図", bold_style), Paragraph(f"口絵:{data.get('field-65', '')}(丁:{data.get('field-66', '')}) / 折図:{data.get('field-69', '')}", normal_style),
         Paragraph("半製品預かり", bold_style), Paragraph(str(data.get('field-81', '')), normal_style)],
    ]
    spec_table = Table(spec_data, colWidths=[25*mm, 70*mm, 25*mm, 70*mm])
    spec_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
        ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(spec_table)
    story.append(Spacer(1, 3 * mm))

    # 4. 特記事項・申し送り事項
    note_data = [
        [Paragraph("特記事項 / 申し送り", bold_style)],
        [Paragraph(data.get('field-45', '特記事項なし'), normal_style)]
    ]
    note_table = Table(note_data, colWidths=[190*mm])
    note_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('PADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(note_table)
    
    doc.build(story)
    print(f"製本進行カードPDFを生成しました: {output_filename}")

if __name__ == "__main__":
    # テスト用データ（実際には @pocket から取得したフィールドデータを使用）
    sample_data = {
        "field-2": "261333",
        "field-52": "3月上旬",
        "field-54": "見本 500部",
        "field-56": "配本 0部",
        "field-13": "法とジャーナリズム〈第4版〉",
        "field-8": "勁草書房",
        "field-93": "403943",
        "field-14": "三秀舎",
        "field-19": "教科書",
        "field-28": "3",
        "field-29": "1",
        "field-20": "3,000",
        "field-3": "3,000",
        "field-4": "500",
        "field-82": "448",
        "field-41": "アジロ無線",
        "field-42": "三方",
        "field-43": "小口糊",
        "field-44": "左開き",
        "field-24": "210.0",
        "field-25": "148.0",
        "field-49": "表紙 カバー 帯 扉",
        "field-48": "淡色白孔雀キク/T76.5kg",
        "field-65": "無",
        "field-66": "",
        "field-69": "無",
        "field-81": "0",
        "field-45": "袋払い 3mm / 教科書用・小口糊, 左, 扉, カバーマットPP加工"
    }
    generate_seihon_card(sample_data)