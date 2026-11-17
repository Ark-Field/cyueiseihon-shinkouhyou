import os
import io
import requests
import pandas as pd
from datetime import datetime
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# ページ基本設定
# ==========================================
st.set_page_config(page_title="製本帳票ツール（進行カード・面掛表）", layout="centered")
st.title("🖨️ 製本帳票（進行カード・面掛表）")

# ==========================================
# @pocket認証設定 (Secretsから取得)
# ==========================================
if "pocket" in st.secrets:
    POCKET_API_KEY = st.secrets["pocket"].get("api_key", "XZtsg3OrDFOh0VayOI83j84Oc3YIzm7oujbI0T08")
    POCKET_BASE_URL = st.secrets["pocket"].get("base_url", "https://app070.at-pocket.com/arkf05_bb")
else:
    POCKET_API_KEY = "XZtsg3OrDFOh0VayOI83j84Oc3YIzm7oujbI0T08"
    POCKET_BASE_URL = "https://app070.at-pocket.com/arkf05_bb"

APP_ID = 34

# ==========================================
# 日本語フォント登録
# ==========================================
FONT_PATH = "NotoSansJP-Regular.ttf"
if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont('Japanese', FONT_PATH))
    DEFAULT_FONT = 'Japanese'
else:
    DEFAULT_FONT = 'Helvetica'

# ==========================================
# APIからデータ取得関数
# ==========================================
def fetch_record_by_order(order_no):
    url = f"{POCKET_BASE_URL}/api/apps/{APP_ID}/records"
    headers = {"X-At-Pocket-API-Key": POCKET_API_KEY, "Accept": "application/json"}
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
        st.error(f"API通信エラー: {e}")
    return None

# ==========================================
# 1. 製本進行カード PDF生成
# ==========================================
def generate_progress_card_pdf(data):
    pdf_path = "temp_progress_card.pdf"
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=8 * mm, leftMargin=8 * mm,
        topMargin=6 * mm, bottomMargin=6 * mm
    )
    
    title_style = ParagraphStyle('CardTitle', fontName=DEFAULT_FONT, fontSize=14, leading=18, alignment=1)
    normal_style = ParagraphStyle('CardNormal', fontName=DEFAULT_FONT, fontSize=8.5, leading=11)
    bold_style = ParagraphStyle('CardBold', fontName=DEFAULT_FONT, fontSize=8.5, leading=11)
    center_style = ParagraphStyle('CardCenter', fontName=DEFAULT_FONT, fontSize=8.5, leading=10, alignment=1)
    
    spec_val_style = ParagraphStyle('SpecVal10', fontName=DEFAULT_FONT, fontSize=10, leading=13)

    p32_1 = data.get('field-30', '')
    p16 = data.get('field-31', '')
    p8 = data.get('field-32', '')
    p4 = data.get('field-33', '')
    p32_2 = data.get('field-34', '')
    total_pages_str = f"32P:{p32_1} / 16P:{p16} / 8P:{p8} / 4P:{p4} / 32P:{p32_2}"

    story = []

    # ヘッダー
    space_4 = "&nbsp;&nbsp;&nbsp;&nbsp;"
    ban_value = data.get('field-27', '')
    ban_str = f"<b>版:</b>{space_4}{ban_value}"

    header_data = [
        [
            Paragraph(ban_str, normal_style),
            Paragraph(f"<b>納期:</b> {data.get('field-52', '')}", normal_style),
            Paragraph(f"<b>見本:</b> {data.get('field-54', '')}", normal_style),
            Paragraph(f"<b>配本:</b> {data.get('field-56', '')}", normal_style),
            Paragraph("製本進行カード", title_style),
            Paragraph(f"<b>受注番号:</b> {data.get('field-2', '')}", bold_style)
        ]
    ]
    header_table = Table(header_data, colWidths=[24*mm, 30*mm, 25*mm, 25*mm, 52*mm, 38*mm], rowHeights=[9*mm])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (4,0), (4,0), colors.lightgrey),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 1.5 * mm))

    # 基本情報
    info_data = [
        [Paragraph("書名", bold_style), Paragraph(str(data.get('field-13', '')), normal_style),
         Paragraph("出版社", bold_style), Paragraph(str(data.get('field-8', '')), normal_style)],
        [Paragraph("ISBN", bold_style), Paragraph(str(data.get('field-93', '')), normal_style),
         Paragraph("印刷所", bold_style), Paragraph(str(data.get('field-14', '')), normal_style)],
        [Paragraph("種別/刷数", bold_style), Paragraph(f"{data.get('field-19', '')} / {data.get('field-28', '')}刷 (回:{data.get('field-29', '')})", normal_style),
         Paragraph("定価", bold_style), Paragraph(f"¥{data.get('field-20', '')}", normal_style)],
        [Paragraph("総頁数", bold_style), Paragraph(total_pages_str, normal_style),
         Paragraph("印刷/製本数", bold_style), Paragraph(f"印:{data.get('field-3', '')}部 / 製:{data.get('field-4', '')}部", normal_style)],
    ]
    info_table = Table(info_data, colWidths=[22*mm, 72*mm, 22*mm, 78*mm], rowHeights=[6.5*mm, 6.5*mm, 6.5*mm, 6.5*mm])
    info_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 1.5 * mm))

    # 仕様情報
    f_val = data.get('field-49', '')
    m_val = data.get('field-48', '')
    e_val = f"口絵:{data.get('field-65', '')}(丁:{data.get('field-66', '')}) / 折図:{data.get('field-69', '')}"
    h_val = str(data.get('field-81', ''))

    spec_data = [
        [Paragraph("仕上仕様", bold_style), Paragraph(f"{data.get('field-41', '')} / {data.get('field-42', '')}", normal_style),
         Paragraph("背固仕様", bold_style), Paragraph(str(data.get('field-43', '')), normal_style)],
        [Paragraph("仕上寸法", bold_style), Paragraph(f"天地:{data.get('field-24', '')} / 左右:{data.get('field-25', '')} ({data.get('field-44', '')})", spec_val_style),
         Paragraph("付物", bold_style), Paragraph(f"□ {f_val}", spec_val_style)],
        [Paragraph("口絵 / 折図", bold_style), Paragraph(e_val, normal_style),
         Paragraph("見返し", bold_style), Paragraph(f"□ {m_val}", spec_val_style)],
        [Paragraph("半製品預かり", bold_style), Paragraph(h_val, normal_style),
         Paragraph("", bold_style), Paragraph("", bold_style)],
    ]
    spec_table = Table(spec_data, colWidths=[22*mm, 72*mm, 22*mm, 78*mm], rowHeights=[6.5*mm, 15*mm, 7*mm, 6.5*mm])
    spec_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
        ('SPAN', (0,3), (1,3)),
        ('SPAN', (2,3), (3,3)),
    ]))
    story.append(spec_table)
    story.append(Spacer(1, 1.5 * mm))

    # 数字マス目
    row1_nums = [Paragraph(str(i), center_style) for i in range(1, 19)]
    row2_nums = [Paragraph(str(i), center_style) for i in range(19, 37)]
    row3_nums = [Paragraph(str(i), center_style) for i in range(37, 55)]
    
    grid_data = [row1_nums, row2_nums, row3_nums]
    col_widths = [10.7 * mm] * 18
    num_grid_table = Table(grid_data, colWidths=col_widths, rowHeights=[11.5*mm, 11.5*mm, 11.5*mm])
    num_grid_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
        ('BACKGROUND', (0,0), (-1,-1), colors.white),
    ]))
    story.append(num_grid_table)
    story.append(Spacer(1, 1.5 * mm))

    # 特記事項欄
    note_data = [
        [Paragraph("特記事項 / 申し送り欄（フリー手書きスペース）", bold_style)],
        [Paragraph(str(data.get('field-45', '特記事項なし')), normal_style)]
    ]
    note_table = Table(note_data, colWidths=[194*mm], rowHeights=[6.5*mm, 28*mm])
    note_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('VALIGN', (0,1), (0,1), 'TOP'),
        ('VALIGN', (0,0), (0,0), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(note_table)
    story.append(Spacer(1, 1.5 * mm))

    # ■工程チェックリスト
    chk_header = Paragraph("<b>■ 工程チェックリスト</b>", bold_style)
    story.append(chk_header)
    story.append(Spacer(1, 1 * mm))

    space_4_chars = "&nbsp;&nbsp;&nbsp;&nbsp;"
    date_seal_str = f"<b>作業日:{space_4_chars}年{space_4_chars}月{space_4_chars}日{space_4_chars}■担当印&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</b>"
    chk_content_bold_style = ParagraphStyle('ChkContentBold75', fontName=DEFAULT_FONT, fontSize=7.5, leading=10)

    checklist_data = [
        [Paragraph("<b>【丁合】</b>", bold_style), Paragraph("<b>□ ・背丁、ページの確認</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("", bold_style), Paragraph("<b>□ ・書名・台数・本数の確認</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("", bold_style), Paragraph("<b>□ ・作業終了後余り丁・表紙の整理</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("<b>【無線】</b>", bold_style), Paragraph("<b>□ ・表紙(書名、巻数、天グル防止)の確認</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("", bold_style), Paragraph("<b>□ ・見返し・扉・口絵等の確認</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("", bold_style), Paragraph("<b>□ ・カット・アジロ・小口糊の確認</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("<b>【三方】</b>", bold_style), Paragraph("<b>□ ・寸法・落としの確認</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("", bold_style), Paragraph("<b>□ ・アメ・カブリ・字切れ・トンボ残りの確認</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("", bold_style), Paragraph("<b>□ ・折込・汚れの確認</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("<b>【仕上げ】</b>", bold_style), Paragraph("<b>□ ・背丁・ページの確認 / 付物等書名・ISBN・定価等の確認</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("", bold_style), Paragraph("<b>□ ・付物の印刷ヤレ確認 / 梱包の糊の具合・判子の状態の確認</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
        [Paragraph("", bold_style), Paragraph("<b>□ ・作業後の片付けの確認 / 製品の最終チェック</b>", chk_content_bold_style), Paragraph(date_seal_str, chk_content_bold_style)],
    ]

    chk_table = Table(checklist_data, colWidths=[22*mm, 112*mm, 60*mm], rowHeights=[9*mm] * 12)
    chk_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 1),
        ('SPAN', (0,0), (0,2)),
        ('SPAN', (0,3), (0,5)),
        ('SPAN', (0,6), (0,8)),
        ('SPAN', (0,9), (0,11)),
    ]))
    story.append(chk_table)

    doc.build(story)
    with open(pdf_path, "rb") as f:
        return f.read()

# ==========================================
# 2. 面掛表・貼り込み表 PDF生成（2ページ構成）
# ==========================================
def generate_menkake_pdf(data):
    pdf_path = "temp_menkake.pdf"
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=8 * mm, leftMargin=8 * mm,
        topMargin=4 * mm, bottomMargin=4 * mm
    )
    
    title_style = ParagraphStyle('Title16', fontName=DEFAULT_FONT, fontSize=16, leading=20, alignment=1)
    normal_style = ParagraphStyle('Normal', fontName=DEFAULT_FONT, fontSize=7.5, leading=10)
    bold_style = ParagraphStyle('Bold', fontName=DEFAULT_FONT, fontSize=7.5, leading=10)
    center_style = ParagraphStyle('Center', fontName=DEFAULT_FONT, fontSize=7.5, leading=10, alignment=1)
    right_style = ParagraphStyle('Right', fontName=DEFAULT_FONT, fontSize=7.5, leading=10, alignment=2)

    center_7pt = ParagraphStyle('Center7Pt', fontName=DEFAULT_FONT, fontSize=7, leading=9, alignment=1)
    bold_center_7pt = ParagraphStyle('BoldCenter7Pt', fontName=DEFAULT_FONT, fontSize=7, leading=9, alignment=1)

    story = []

    # 1ページ目：面掛表 ＆ 表2
    story.append(Paragraph("<b>面掛表</b>", title_style))
    story.append(Spacer(1, 1 * mm))

    top_info_data = [
        [
            Paragraph(f"<b>受注番号:</b> {data.get('field-2', '')}", normal_style),
            Paragraph(f"<b>ISBN:</b> {data.get('field-93', '')}", normal_style)
        ]
    ]
    top_info_table = Table(top_info_data, colWidths=[97*mm, 97*mm], rowHeights=[5*mm])
    top_info_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(top_info_table)
    story.append(Spacer(1, 1 * mm))

    title_blank = " " * 55
    book_title_data = [[Paragraph(f"<b>書名:</b> {title_blank}", normal_style)]]
    book_title_table = Table(book_title_data, colWidths=[194*mm], rowHeights=[5*mm])
    book_title_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(book_title_table)
    story.append(Spacer(1, 1 * mm))

    pasting_headers = [
        Paragraph("<b>面掛</b>", bold_style),
        Paragraph("<b>摘要</b>", bold_style),
        Paragraph("<b>面掛</b>", bold_style),
        Paragraph("<b>摘要</b>", bold_style),
        Paragraph("<b>備考</b>", bold_style)
    ]
    pasting_rows = [pasting_headers]
    menkake_cell_text = " " * 5 + "×" + " " * 5
    for i in range(10):
        pasting_rows.append([
            Paragraph(menkake_cell_text, center_style),
            Paragraph("", normal_style),
            Paragraph(menkake_cell_text, center_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
        ])

    pasting_table = Table(
        pasting_rows,
        colWidths=[38*mm, 22*mm, 38*mm, 22*mm, 74*mm],
        rowHeights=[6.5*mm] * 11
    )
    pasting_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('PADDING', (0,0), (-1,-1), 1),
    ]))
    story.append(pasting_table)
    story.append(Spacer(1, 1 * mm))

    note_blank = " " * 55
    note_data = [[Paragraph(f"<b>注意事項:</b> {note_blank}", normal_style)]]
    note_table = Table(note_data, colWidths=[194*mm], rowHeights=[30*mm])
    note_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(note_table)
    story.append(Spacer(1, 4 * mm))

    space_25cm = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    work_date_str = f"<b>作業日:</b>{space_25cm}年{space_25cm}月{space_25cm}日"

    work_date_data = [[Paragraph(work_date_str, normal_style)]]
    work_date_table = Table(work_date_data, colWidths=[194*mm], rowHeights=[5*mm])
    work_date_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(work_date_table)
    story.append(Spacer(1, 1 * mm))

    sub_headers = [
        Paragraph("<b>No</b>", bold_center_7pt), Paragraph("<b>担当</b>", center_style), Paragraph("<b>折メーター</b>", center_style), Paragraph("<b>折損</b>", bold_center_7pt),
        Paragraph("<b>No</b>", bold_center_7pt), Paragraph("<b>担当</b>", center_style), Paragraph("<b>折メーター</b>", center_style), Paragraph("<b>折損</b>", bold_center_7pt),
        Paragraph("<b>No</b>", bold_center_7pt), Paragraph("<b>担当</b>", center_style), Paragraph("<b>折メーター</b>", center_style), Paragraph("<b>折損</b>", bold_center_7pt),
        Paragraph("<b>No</b>", bold_center_7pt), Paragraph("<b>担当</b>", center_style), Paragraph("<b>折メーター</b>", center_style), Paragraph("<b>折損</b>", bold_center_7pt)
    ]
    sub_rows = [sub_headers]
    for r in range(10):
        sub_rows.append([
            Paragraph(str(r + 1), center_7pt), Paragraph("", normal_style), Paragraph("", normal_style), Paragraph("", center_7pt),
            Paragraph(str(r + 11), center_7pt), Paragraph("", normal_style), Paragraph("", normal_style), Paragraph("", center_7pt),
            Paragraph(str(r + 21), center_7pt), Paragraph("", normal_style), Paragraph("", normal_style), Paragraph("", center_7pt),
            Paragraph(str(r + 31), center_7pt), Paragraph("", normal_style), Paragraph("", normal_style), Paragraph("", center_7pt),
        ])
    
    sub_col_widths = [7.5*mm, 13*mm, 19*mm, 9*mm] * 4
    sub_table = Table(sub_rows, colWidths=sub_col_widths, rowHeights=[5.2*mm] * 11)
    sub_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('PADDING', (0,0), (-1,-1), 1),
    ]))
    story.append(sub_table)
    story.append(Spacer(1, 1 * mm))

    note_data2 = [[Paragraph(f"<b>注意事項:</b> {note_blank}", normal_style)]]
    note_table2 = Table(note_data2, colWidths=[194*mm], rowHeights=[25*mm])
    note_table2.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(note_table2)

    # 2ページ目：貼り込み表
    story.append(PageBreak())
    story.append(Paragraph("<b>貼り込み表</b>", title_style))
    story.append(Spacer(1, 2 * mm))

    work_date_pasting = f"<b>作業日:</b>{space_25cm}年{space_25cm}月{space_25cm}日"
    pasting_info_data = [
        [
            Paragraph("<b>書名:</b>               ", normal_style),
            Paragraph(f"<b>部数:</b> {data.get('field-4', '500')} 部", normal_style),
            Paragraph(work_date_pasting, normal_style)
        ]
    ]
    pasting_info_table = Table(pasting_info_data, colWidths=[84*mm, 45*mm, 65*mm], rowHeights=[6*mm])
    pasting_info_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(pasting_info_table)
    story.append(Spacer(1, 2 * mm))

    pasting_headers2 = [
        Paragraph("<b>貼り込み箇所</b>", bold_style),
        Paragraph("<b>本数</b>", bold_style),
        Paragraph("<b>半端</b>", bold_style),
        Paragraph("<b>ヤレ枚数</b>", bold_style),
        Paragraph("<b>備考</b>", bold_style)
    ]
    pasting_rows2 = [pasting_headers2]
    nbsp18 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    pasting_cell_text = f"{nbsp18}折〜{nbsp18}折"
    
    for i in range(24):
        pasting_rows2.append([
            Paragraph(pasting_cell_text, normal_style),
            Paragraph("本", right_style),
            Paragraph("本", right_style),
            Paragraph("枚", right_style),
            Paragraph("", normal_style),
        ])

    pasting_table2 = Table(
        pasting_rows2, 
        colWidths=[60*mm, 24.5*mm, 24.5*mm, 24.5*mm, 60.5*mm], 
        rowHeights=[9.2*mm] * 25
    )
    pasting_table2.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (1,1), (3,-1), 'RIGHT'),
        ('PADDING', (0,0), (-1,-1), 2),
    ]))
    story.append(pasting_table2)

    doc.build(story)
    with open(pdf_path, "rb") as f:
        return f.read()

# ==========================================
# 3. Excelワークブック生成関数（製本進行カード）
# ==========================================
def generate_excel_workbook(data):
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "製本進行カード"
    ws1.views.sheetView[0].showGridLines = True

    font_name = "Meiryo"
    title_font = Font(name=font_name, size=14, bold=True)
    bold_font = Font(name=font_name, size=9, bold=True)
    normal_font = Font(name=font_name, size=9)

    gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    accent_fill = PatternFill(start_color="E0E0E0", end_color="E0E0E0", fill_type="solid")
    thin_side = Side(style='thin', color='000000')
    border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    for i in range(1, 19):
        ws1.column_dimensions[get_column_letter(i)].width = 4.2

    # Row 1: Header
    ws1.merge_cells('A1:B1'); ws1['A1'] = f"版: {data.get('field-27', '')}"
    ws1.merge_cells('C1:D1'); ws1['C1'] = f"納期: {data.get('field-52', '')}"
    ws1.merge_cells('E1:F1'); ws1['E1'] = f"見本: {data.get('field-54', '')}"
    ws1.merge_cells('G1:H1'); ws1['G1'] = f"配本: {data.get('field-56', '')}"
    ws1.merge_cells('I1:N1'); ws1['I1'] = "製本進行カード"
    ws1['I1'].alignment = Alignment(horizontal='center', vertical='center')
    ws1['I1'].font = title_font
    ws1['I1'].fill = accent_fill
    ws1.merge_cells('O1:R1'); ws1['O1'] = f"受注番号: {data.get('field-2', '')}"

    for col in range(1, 19):
        cell = ws1.cell(row=1, column=col)
        cell.border = border
        if col not in range(9, 15):
            cell.font = bold_font
            cell.alignment = Alignment(horizontal='left', vertical='center')
    ws1.row_dimensions[1].height = 28

    p32_1 = data.get('field-30', '')
    p16 = data.get('field-31', '')
    p8 = data.get('field-32', '')
    p4 = data.get('field-33', '')
    p32_2 = data.get('field-34', '')
    total_pages_str = f"32P:{p32_1} / 16P:{p16} / 8P:{p8} / 4P:{p4} / 32P:{p32_2}"

    basic_info = [
        ("書名", str(data.get('field-13', '')), "出版社", str(data.get('field-8', ''))),
        ("ISBN", str(data.get('field-93', '')), "印刷所", str(data.get('field-14', ''))),
        ("種別/刷数", f"{data.get('field-19', '')} / {data.get('field-28', '')}刷", "定価", f"¥{data.get('field-20', '')}"),
        ("総頁数", total_pages_str, "印刷/製本数", f"印:{data.get('field-3', '')}部 / 製:{data.get('field-4', '')}部")
    ]

    for idx, row_data in enumerate(basic_info, start=2):
        ws1.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=2)
        ws1.merge_cells(start_row=idx, start_column=3, end_row=idx, end_column=10)
        ws1.merge_cells(start_row=idx, start_column=11, end_row=idx, end_column=12)
        ws1.merge_cells(start_row=idx, start_column=13, end_row=idx, end_column=18)
        
        ws1.cell(row=idx, column=1, value=row_data[0])
        ws1.cell(row=idx, column=3, value=row_data[1])
        ws1.cell(row=idx, column=11, value=row_data[2])
        ws1.cell(row=idx, column=13, value=row_data[3])
        
        for c in range(1, 19):
            cell = ws1.cell(row=idx, column=c)
            cell.border = border
            cell.font = normal_font if c in [3, 13] else bold_font
            if c in [1, 11]:
                cell.fill = gray_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')
        ws1.row_dimensions[idx].height = 20

    spec_info = [
        ("仕上仕様", f"{data.get('field-41', '')} / {data.get('field-42', '')}", "背固仕様", str(data.get('field-43', ''))),
        ("仕上寸法", f"天地:{data.get('field-24', '')} / 左右:{data.get('field-25', '')}", "付物", f"□ {data.get('field-49', '')}"),
        ("口絵 / 折図", f"口絵:{data.get('field-65', '')} / 折図:{data.get('field-69', '')}", "見返し", f"□ {data.get('field-48', '')}"),
        ("半製品預かり", str(data.get('field-81', '')), "", "")
    ]

    for idx, row_data in enumerate(spec_info, start=6):
        ws1.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=2)
        ws1.merge_cells(start_row=idx, start_column=3, end_row=idx, end_column=10)
        ws1.merge_cells(start_row=idx, start_column=11, end_row=idx, end_column=12)
        ws1.merge_cells(start_row=idx, start_column=13, end_row=idx, end_column=18)
        
        ws1.cell(row=idx, column=1, value=row_data[0])
        ws1.cell(row=idx, column=3, value=row_data[1])
        ws1.cell(row=idx, column=11, value=row_data[2])
        ws1.cell(row=idx, column=13, value=row_data[3])
        
        for c in range(1, 19):
            cell = ws1.cell(row=idx, column=c)
            cell.border = border
            cell.font = normal_font if c in [3, 13] else bold_font
            if c in [1, 11]:
                cell.fill = gray_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')
                
        if idx == 6: ws1.row_dimensions[idx].height = 20
        elif idx == 7: ws1.row_dimensions[idx].height = 42
        elif idx in [8, 9]: ws1.row_dimensions[idx].height = 20

    ws1.merge_cells('C9:H9')
    ws1.merge_cells('M9:R9')

    page_nums = [list(range(1, 19)), list(range(19, 37)), list(range(37, 55))]
    for r_idx, nums in enumerate(page_nums, start=10):
        ws1.row_dimensions[r_idx].height = 30
        for c_idx, num in enumerate(nums, start=1):
            cell = ws1.cell(row=r_idx, column=c_idx, value=num)
            cell.border = border
            cell.font = bold_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

    ws1.merge_cells('A13:R13'); ws1['A13'] = "特記事項 / 申し送り欄（フリー手書きスペース）"
    ws1['A13'].font = bold_font; ws1['A13'].fill = gray_fill
    ws1['A13'].alignment = Alignment(horizontal='left', vertical='center')
    for c in range(1, 19): ws1.cell(row=13, column=c).border = border
    ws1.row_dimensions[13].height = 20

    for r in range(14, 18):
        ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=18)
        for c in range(1, 19): ws1.cell(row=r, column=c).border = border
        ws1.row_dimensions[r].height = 20
    
    if data.get('field-45'):
        ws1['A14'] = str(data.get('field-45'))

    ws1.merge_cells('A18:R18'); ws1['A18'] = "■ 工程チェックリスト"
    ws1['A18'].font = bold_font
    for c in range(1, 19): ws1.cell(row=18, column=c).border = border
    ws1.row_dimensions[18].height = 20

    processes = [
        ("【丁合】", "□ ・背丁、ページの確認"),
        ("【丁合】", "□ ・書名・台数・本数の確認"),
        ("【丁合】", "□ ・作業終了後余り丁・表紙の整理"),
        ("【無線】", "□ ・表紙(書名、巻数、天グル防止)の確認"),
        ("【無線】", "□ ・見返し・扉・口絵等の確認"),
        ("【無線】", "□ ・カット・アジロ・小口糊の確認"),
        ("【三方】", "□ ・寸法・落としの確認"),
        ("【三方】", "□ ・アメ・カブリ・字切れ・トンボ残りの確認"),
        ("【三方】", "□ ・折込・汚れの確認"),
        ("【仕上げ】", "□ ・背丁・ページの確認 / 付物等書名・ISBN・定価等の確認"),
        ("【仕上げ】", "□ ・付物の印刷ヤレ確認 / 梱包の糊の具合・判子の状態の確認"),
        ("【仕上げ】", "□ ・作業後の片付けの確認 / 製品の最終チェック"),
    ]

    for idx, (proc, desc) in enumerate(processes, start=19):
        ws1.row_dimensions[idx].height = 25
        ws1.merge_cells(start_row=idx, start_column=1, end_row=idx, end_column=2)
        ws1.merge_cells(start_row=idx, start_column=3, end_row=idx, end_column=11)
        ws1.merge_cells(start_row=idx, start_column=12, end_row=idx, end_column=18)
        
        ws1.cell(row=idx, column=1, value=proc)
        ws1.cell(row=idx, column=3, value=desc)
        ws1.cell(row=idx, column=12, value="作業日:    年    月    日    ■担当印")
        
        for c in range(1, 19):
            cell = ws1.cell(row=idx, column=c)
            cell.border = border
            cell.font = bold_font
            if c <= 2:
                cell.fill = gray_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif c >= 12:
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.font = normal_font
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center')

    ws1.merge_cells('A19:B21'); ws1.merge_cells('A22:B24'); ws1.merge_cells('A25:B27'); ws1.merge_cells('A28:B30')

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# ==========================================
# メイン画面処理 (URLパラメータ連動 ＆ ボタン選択)
# ==========================================
query_params = st.query_params
url_order_no = query_params.get("ord", None)

st.info("💡 @pocketの受注番号から、必要な帳票（製本進行カード または 面掛表・貼り込み表）を選んでPDFまたはExcel出力できます。")

input_order_no = st.text_input("受注番号を入力", value=url_order_no if url_order_no else "")

if input_order_no:
    with st.spinner(f"受注番号 {input_order_no} のデータを@pocketから取得中..."):
        record_data = fetch_record_by_order(input_order_no)
        
    if record_data:
        st.success(f"✨ 対象レコードが見つかりました: **{record_data.get('field-13', '')}** ({record_data.get('field-8', '')})")
        
        st.markdown("---")
        st.subheader("🖨️ 出力する帳票と形式を選択してください")
        
        col_btn1, col_btn2 = st.columns(2)
        
        # 1. 製本進行カード出力ボタン（PDF & Excel）
        with col_btn1:
            st.markdown("#### 📋 製本進行カード")
            st.write("数字マス目を広げたチェックリスト付き")
            
            progress_pdf_bytes = generate_progress_card_pdf(record_data)
            st.download_button(
                label="📥 進行カードPDFをダウンロード",
                data=progress_pdf_bytes,
                file_name=f"製本進行カード_{input_order_no}.pdf",
                mime="application/pdf",
                type="primary",
                key="btn_progress_pdf"
            )
            
            progress_excel_bytes = generate_excel_workbook(record_data)
            st.download_button(
                label="📊 進行カードExcelをダウンロード",
                data=progress_excel_bytes,
                file_name=f"製本進行カード_{input_order_no}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_progress_excel"
            )
            
        # 2. 面掛表・貼り込み表出力ボタン（PDF）
        with col_btn2:
            st.markdown("#### 📋 面掛表・貼り込み表")
            st.write("面掛表（1頁目）＆ 貼り込み表（2頁目）の2頁構成")
            
            menkake_pdf_bytes = generate_menkake_pdf(record_data)
            st.download_button(
                label="📥 面掛表・貼り込み表PDFをダウンロード",
                data=menkake_pdf_bytes,
                file_name=f"面掛表_貼り込み表_{input_order_no}.pdf",
                mime="application/pdf",
                type="primary",
                key="btn_menkake"
            )
    else:
        st.error(f"❌ 受注番号「{input_order_no}」に一致するレコードが見つかりませんでした。")
else:
    st.warning("⚠️ 受注番号を入力してください（URLパラメータ `?ord=番号` からも自動読み込みされます）。")