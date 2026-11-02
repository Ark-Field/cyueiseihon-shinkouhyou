import os
import requests
import pandas as pd
from datetime import datetime
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ==========================================
# ページ基本設定
# ==========================================
st.set_page_config(page_title="製本進行カード自動生成ツール", layout="centered")
st.title("🖨️ 製本進行カード自動生成ツール")

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
# PDF生成ロジック（チェックリスト統合版）
# ==========================================
def generate_pdf_bytes(data):
    pdf_path = "temp_card.pdf"
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=8 * mm, leftMargin=8 * mm,
        topMargin=8 * mm, bottomMargin=8 * mm
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('CardTitle', fontName=DEFAULT_FONT, fontSize=14, leading=18, alignment=1)
    normal_style = ParagraphStyle('CardNormal', fontName=DEFAULT_FONT, fontSize=7.5, leading=9.5)
    bold_style = ParagraphStyle('CardBold', fontName=DEFAULT_FONT, fontSize=7.5, leading=9.5)

    story = []

    # 1. ヘッダー
    header_data = [
        [
            Paragraph(f"<b>納期:</b> {data.get('field-52', '')}", normal_style),
            Paragraph(f"<b>見本:</b> {data.get('field-54', '')}", normal_style),
            Paragraph(f"<b>配本:</b> {data.get('field-56', '')}", normal_style),
            Paragraph("製本進行カード", title_style),
            Paragraph(f"<b>受注番号:</b> {data.get('field-2', '')}", bold_style)
        ]
    ]
    header_table = Table(header_data, colWidths=[38*mm, 30*mm, 30*mm, 58*mm, 38*mm])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (3,0), (3,0), colors.lightgrey),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 2 * mm))

    # 2. 基本情報
    info_data = [
        [Paragraph("書名", bold_style), Paragraph(str(data.get('field-13', '')), normal_style),
         Paragraph("出版社", bold_style), Paragraph(str(data.get('field-8', '')), normal_style)],
        [Paragraph("ISBN", bold_style), Paragraph(str(data.get('field-93', '')), normal_style),
         Paragraph("印刷所", bold_style), Paragraph(str(data.get('field-14', '')), normal_style)],
        [Paragraph("種別/刷数", bold_style), Paragraph(f"{data.get('field-19', '')} / {data.get('field-28', '')}刷 (回:{data.get('field-29', '')})", normal_style),
         Paragraph("定価", bold_style), Paragraph(f"¥{data.get('field-20', '')}", normal_style)],
        [Paragraph("総頁数", bold_style), Paragraph(f"{data.get('field-82', '')} 頁", normal_style),
         Paragraph("印刷/製本数", bold_style), Paragraph(f"印:{data.get('field-3', '')}部 / 製:{data.get('field-4', '')}部", normal_style)],
    ]
    info_table = Table(info_data, colWidths=[22*mm, 72*mm, 22*mm, 78*mm])
    info_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
        ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2.5),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 2 * mm))

    # 3. 仕様情報
    spec_data = [
        [Paragraph("仕上仕様", bold_style), Paragraph(f"{data.get('field-41', '')} / {data.get('field-42', '')}", normal_style),
         Paragraph("背固仕様", bold_style), Paragraph(str(data.get('field-43', '')), normal_style)],
        [Paragraph("仕上寸法", bold_style), Paragraph(f"天地:{data.get('field-24', '')} / 左右:{data.get('field-25', '')} ({data.get('field-44', '')})", normal_style),
         Paragraph("付物/見返し", bold_style), Paragraph(f"付:{data.get('field-49', '')} / 見:{data.get('field-48', '')}", normal_style)],
        [Paragraph("口絵 / 折図", bold_style), Paragraph(f"口絵:{data.get('field-65', '')}(丁:{data.get('field-66', '')}) / 折図:{data.get('field-69', '')}", normal_style),
         Paragraph("半製品預かり", bold_style), Paragraph(str(data.get('field-81', '')), normal_style)],
    ]
    spec_table = Table(spec_data, colWidths=[22*mm, 72*mm, 22*mm, 78*mm])
    spec_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
        ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2.5),
    ]))
    story.append(spec_table)
    story.append(Spacer(1, 2 * mm))

    # 4. 特記事項
    note_data = [
        [Paragraph("特記事項 / 申し送り", bold_style)],
        [Paragraph(str(data.get('field-45', '特記事項なし')), normal_style)]
    ]
    note_table = Table(note_data, colWidths=[194*mm])
    note_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(note_table)
    story.append(Spacer(1, 2 * mm))

    # 5. チェックリスト項目（元画像に基づく追加）[cite: 2]
    chk_header = Paragraph("<b>■ 工程チェックリスト</b>", bold_style)
    story.append(chk_header)
    story.append(Spacer(1, 1 * mm))

    checklist_data = [
        [Paragraph("<b>【丁合】</b>[cite: 2]", bold_style), Paragraph("・背丁、ページの確認[cite: 2]<br/>・書名・台数・本数の確認[cite: 2]<br/>・作業終了後余り丁・表紙の整理[cite: 2]", normal_style), Paragraph("作業日:  年 月 日<br/>担当者印:[cite: 2]", normal_style)],
        [Paragraph("<b>【無線】</b>[cite: 2]", bold_style), Paragraph("・表紙(書名、巻数、天グル防止)の確認[cite: 2]<br/>・見返し・扉・口絵等の確認[cite: 2]<br/>・カット・アジロ・小口糊の確認[cite: 2]", normal_style), Paragraph("作業日:  年 月 日<br/>担当者印:[cite: 2]", normal_style)],
        [Paragraph("<b>【三方】</b>[cite: 2]", bold_style), Paragraph("・寸法・落としの確認[cite: 2]<br/>・アメ・カブリ・字切れ・トンボ残りの確認[cite: 2]<br/>・折込・汚れの確認[cite: 2]", normal_style), Paragraph("作業日:  年 月 日<br/>担当者印:[cite: 2]", normal_style)],
        [Paragraph("<b>【仕上げ】</b>[cite: 2]", bold_style), Paragraph("・背丁・ページの確認 / 付物等書名・ISBN・定価等の確認[cite: 2]<br/>・付物の印刷ヤレ確認 / 梱包の糊の具合・判子の状態の確認[cite: 2]<br/>・作業後の片付けの確認 / 製品の最終チェック[cite: 2]", normal_style), Paragraph("作業日:  年 月 日<br/>担当者印:[cite: 2]", normal_style)],
    ]

    chk_table = Table(checklist_data, colWidths=[22*mm, 112*mm, 60*mm])
    chk_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2.5),
    ]))
    story.append(chk_table)

    doc.build(story)
    
    with open(pdf_path, "rb") as f:
        pdf_bytes = f.read()
    return pdf_bytes

# ==========================================
# メイン画面処理 (URLパラメータ連動)
# ==========================================
query_params = st.query_params
url_order_no = query_params.get("ord", None)

st.info("💡 @pocketの受注番号から、チェックリスト付き製本進行カードのPDFをワンクリックで生成します[cite: 2]。")

col1, col2 = st.columns([3, 1])
with col1:
    input_order_no = st.text_input("受注番号を入力", value=url_order_no if url_order_no else "")
with col2:
    st.write("")
    st.write("")
    search_btn = st.button("🔍 カード作成", type="primary")

if url_order_no or search_btn:
    target_ord = url_order_no if url_order_no and not search_btn else input_order_no
    if target_ord:
        with st.spinner(f"受注番号 {target_ord} のデータを@pocketから取得中..."):
            record_data = fetch_record_by_order(target_ord)
            
        if record_data:
            st.success(f"✨ レコードが見つかりました: **{record_data.get('field-13', '')}** ({record_data.get('field-8', '')})")
            
            # PDF生成
            pdf_data = generate_pdf_bytes(record_data)
            
            st.download_button(
                label="📥 チェックリスト付き製本進行カードPDFをダウンロード",
                data=pdf_data,
                file_name=f"製本進行カード_{target_ord}.pdf",
                mime="application/pdf",
                type="primary"
            )
        else:
            st.error(f"❌ 受注番号「{target_ord}」に一致するレコードが見つかりませんでした。")
    else:
        st.warning("⚠️ 受注番号を入力してください。")