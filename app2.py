import os
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

# ==========================================
# ページ基本設定
# ==========================================
st.set_page_config(page_title="面掛表・貼り込み表 自動生成ツール", layout="centered")
st.title("📋 面掛表・貼り込み表 自動生成ツール")

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
# 面掛表・貼り込み表 PDF生成ロジック
# ==========================================
def generate_menkake_pdf(data):
    pdf_path = "temp_menkake.pdf"
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=8 * mm, leftMargin=8 * mm,
        topMargin=8 * mm, bottomMargin=8 * mm
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', fontName=DEFAULT_FONT, fontSize=13, leading=16, alignment=1)
    normal_style = ParagraphStyle('Normal', fontName=DEFAULT_FONT, fontSize=7.5, leading=10)
    bold_style = ParagraphStyle('Bold', fontName=DEFAULT_FONT, fontSize=7.5, leading=10)
    center_style = ParagraphStyle('Center', fontName=DEFAULT_FONT, fontSize=7, leading=9, alignment=1)

    story = []

    # ==========================================
    # 【1ページ目】 面掛表
    # ==========================================
    header_data = [
        [
            Paragraph(f"<b>納期:</b> {data.get('field-52', '')}", normal_style),
            Paragraph(f"<b>見本:</b> {data.get('field-54', '')}", normal_style),
            Paragraph(f"<b>配本:</b> {data.get('field-56', '')}", normal_style),
            Paragraph(f"<b>種別:</b> {data.get('field-19', '')}", normal_style),
            Paragraph(f"<b>判型:</b> {data.get('field-83', 'A5')}", normal_style),
            Paragraph(f"<b>印刷部数:</b> {data.get('field-3', '')}", normal_style),
        ]
    ]
    header_table = Table(header_data, colWidths=[32*mm, 28*mm, 28*mm, 30*mm, 26*mm, 50*mm])
    header_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 2 * mm))

    # タイトル・基本情報部分
    title_info_data = [
        [
            Paragraph("<b>書名</b>", bold_style),
            Paragraph(str(data.get('field-13', '')), normal_style),
            Paragraph("<b>面掛表</b><br/>受注番号: " + str(data.get('field-2', '')), title_style)
        ],
        [
            Paragraph("<b>得意先 / 出版社</b>", bold_style),
            Paragraph(str(data.get('field-8', '')), normal_style),
            Paragraph(f"<b>ISBN:</b> {data.get('field-93', '')}", normal_style)
        ],
        [
            Paragraph("<b>印刷所</b>", bold_style),
            Paragraph(str(data.get('field-14', '')), normal_style),
            Paragraph(f"<b>頁数:</b> {data.get('field-82', '')}頁", normal_style)
        ]
    ]
    title_table = Table(title_info_data, colWidths=[30*mm, 94*mm, 70*mm])
    title_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(title_table)
    story.append(Spacer(1, 2 * mm))

    # 注意事項・背固仕様
    note_data = [
        [Paragraph(f"<b>注意事項:</b> {data.get('field-45', '特記事項なし')} | <b>背固仕様:</b> {data.get('field-43', 'アジロ無線')}", normal_style)]
    ]
    note_table = Table(note_data, colWidths=[194*mm])
    note_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(note_table)
    story.append(Spacer(1, 2 * mm))

    # 1〜40の数字グリッド（面掛確認用）
    grid_rows = []
    for r in range(4):
        row_cells = []
        for c in range(10):
            num = r * 10 + c + 1
            row_cells.append(Paragraph(f"<b>{num}</b><br/><font size=5>折:   損:  </font>", center_style))
        grid_rows.append(row_cells)
    
    menkake_grid = Table(grid_rows, colWidths=[19.4*mm]*10, rowHeights=[10*mm]*4)
    menkake_grid.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BACKGROUND', (0,0), (-1,-1), colors.white),
    ]))
    story.append(menkake_grid)
    
    # ページ分割して2ページ目に「貼り込み表」を出力
    story.append(PageBreak())

    # ==========================================
    # 【2ページ目】 貼り込み表[cite: 6]
    # ==========================================
    story.append(Paragraph("<b>■ 貼り込み表</b>[cite: 6]", title_style))
    story.append(Spacer(1, 2 * mm))

    pasting_info_data = [
        [Paragraph(f"<b>書名:</b> {data.get('field-13', '')}", normal_style),
         Paragraph(f"<b>部数:</b> {data.get('field-4', '500')} 部", normal_style),
         Paragraph("<b>作業日:</b>  年 月 日", normal_style)]
    ]
    pasting_info_table = Table(pasting_info_data, colWidths=[94*mm, 50*mm, 50*mm])
    pasting_info_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(pasting_info_table)
    story.append(Spacer(1, 2 * mm))

    # 貼り込み明細テーブル
    pasting_headers = [Paragraph("<b>貼り込み箇所</b>", bold_style), Paragraph("<b>本数</b>", bold_style), Paragraph("<b>半端</b>", bold_style), Paragraph("<b>ヤレ枚数</b>", bold_style), Paragraph("<b>備考</b>", bold_style)]
    pasting_rows = [pasting_headers]
    
    for i in range(12):  # 記入用行を12行作成
        pasting_rows.append([
            Paragraph(f"折 〜 折", normal_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
        ])

    pasting_table = Table(pasting_rows, colWidths=[64*mm, 30*mm, 30*mm, 30*mm, 40*mm], rowHeights=[7*mm]*13)
    pasting_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.lightgrey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (1,1), (-1,-1), 'CENTER'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(pasting_table)

    doc.build(story)
    
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()
    return pdf_bytes

# ==========================================
# メイン画面処理 (URLパラメータ連動)
# ==========================================
query_params = st.query_params
url_order_no = query_params.get("ord", None)

st.info("💡 @pocketの受注番号から、面掛表および貼り込み表のPDF（2ページ構成）を生成します[cite: 5, 6]。")

col1, col2 = st.columns([3, 1])
with col1:
    input_order_no = st.text_input("受注番号を入力", value=url_order_no if url_order_no else "")
with col2:
    st.write("")
    st.write("")
    search_btn = st.button("🔍 面掛表作成", type="primary")

if url_order_no or search_btn:
    target_ord = url_order_no if url_order_no and not search_btn else input_order_no
    if target_ord:
        with st.spinner(f"受注番号 {target_ord} のデータを@pocketから取得中..."):
            record_data = fetch_record_by_order(target_ord)
            
        if record_data:
            st.success(f"✨ レコードが見つかりました: **{record_data.get('field-13', '')}** ({record_data.get('field-8', '')})")
            
            # PDF生成
            pdf_data = generate_menkake_pdf(record_data)
            
            st.download_button(
                label="📥 面掛表・貼り込み表PDFをダウンロード",
                data=pdf_data,
                file_name=f"面掛表_{target_ord}.pdf",
                mime="application/pdf",
                type="primary"
            )
        else:
            st.error(f"❌ 受注番号「{target_ord}」に一致するレコードが見つかりませんでした。")
    else:
        st.warning("⚠️ 受注番号を入力してください。")