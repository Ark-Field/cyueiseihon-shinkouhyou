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
st.set_page_config(page_title="製本帳票自動生成ツール（進行カード・面掛表）", layout="centered")
st.title("🖨️ 製本帳票自動生成ツール（手書き欄広々・調整版）")

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
# 1. 製本進行カード PDF生成（高さ1.5倍拡充版）
# ==========================================
def generate_progress_card_pdf(data):
    pdf_path = "temp_progress_card.pdf"
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=8 * mm, leftMargin=8 * mm,
        topMargin=6 * mm, bottomMargin=6 * mm
    )
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('CardTitle', fontName=DEFAULT_FONT, fontSize=12, leading=15, alignment=1)
    normal_style = ParagraphStyle('CardNormal', fontName=DEFAULT_FONT, fontSize=7.5, leading=10)
    bold_style = ParagraphStyle('CardBold', fontName=DEFAULT_FONT, fontSize=7.5, leading=10)
    center_style = ParagraphStyle('CardCenter', fontName=DEFAULT_FONT, fontSize=6.5, leading=8, alignment=1)
    chk_content_style = ParagraphStyle('ChkContent', fontName=DEFAULT_FONT, fontSize=7, leading=11)

    story = []

    # ヘッダー
    header_data = [
        [
            Paragraph(f"<b>納期:</b> {data.get('field-52', '')}", normal_style),
            Paragraph(f"<b>見本:</b> {data.get('field-54', '')}", normal_style),
            Paragraph(f"<b>配本:</b> {data.get('field-56', '')}", normal_style),
            Paragraph("製本進行カード", title_style),
            Paragraph(f"<b>受注番号:</b> {data.get('field-2', '')}", bold_style)
        ]
    ]
    header_table = Table(header_data, colWidths=[38*mm, 30*mm, 30*mm, 58*mm, 38*mm], rowHeights=[9*mm])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (3,0), (3,0), colors.lightgrey),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 1.5 * mm))

    # 基本情報（高さを約1.5倍の6mm〜7mmに拡大）
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
    info_table = Table(info_data, colWidths=[22*mm, 72*mm, 22*mm, 78*mm], rowHeights=[6.5*mm, 6.5*mm, 6.5*mm, 6.5*mm])
    info_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
        ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 1.5 * mm))

    # 仕様情報（高さを約1.5倍に拡大）
    spec_data = [
        [Paragraph("仕上仕様", bold_style), Paragraph(f"{data.get('field-41', '')} / {data.get('field-42', '')}", normal_style),
         Paragraph("背固仕様", bold_style), Paragraph(str(data.get('field-43', '')), normal_style)],
        [Paragraph("仕上寸法", bold_style), Paragraph(f"天地:{data.get('field-24', '')} / 左右:{data.get('field-25', '')} ({data.get('field-44', '')})", normal_style),
         Paragraph("付物/見返し", bold_style), Paragraph(f"付:{data.get('field-49', '')} / 見:{data.get('field-48', '')}", normal_style)],
        [Paragraph("口絵 / 折図", bold_style), Paragraph(f"口絵:{data.get('field-65', '')}(丁:{data.get('field-66', '')}) / 折図:{data.get('field-69', '')}", normal_style),
         Paragraph("半製品預かり", bold_style), Paragraph(str(data.get('field-81', '')), normal_style)],
    ]
    spec_table = Table(spec_data, colWidths=[22*mm, 72*mm, 22*mm, 78*mm], rowHeights=[6.5*mm, 6.5*mm, 6.5*mm])
    spec_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(spec_table)
    story.append(Spacer(1, 1.5 * mm))

    # 数字マス目（高さを1.5倍の7mmに調整）
    row1_nums = [Paragraph(str(i), center_style) for i in range(1, 19)]
    row2_nums = [Paragraph(str(i), center_style) for i in range(19, 37)]
    row3_nums = [Paragraph(str(i), center_style) for i in range(37, 55)]
    
    grid_data = [row1_nums, row2_nums, row3_nums]
    col_widths = [10.7 * mm] * 18
    num_grid_table = Table(grid_data, colWidths=col_widths, rowHeights=[7*mm, 7*mm, 7*mm])
    num_grid_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 1),
        ('BACKGROUND', (0,0), (-1,-1), colors.white),
    ]))
    story.append(num_grid_table)
    story.append(Spacer(1, 1.5 * mm))

    # 特記事項欄（高さを約1.5倍に拡大）
    note_data = [
        [Paragraph("特記事項 / 申し送り欄（フリー手書きスペース）[cite: 3]", bold_style)],
        [Paragraph(str(data.get('field-45', '特記事項なし')) + "<br/><br/><br/><br/>", normal_style)]
    ]
    note_table = Table(note_data, colWidths=[194*mm])
    note_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(note_table)
    story.append(Spacer(1, 1.5 * mm))

    # 工程チェックリスト（行高さを1.5倍の約17.5mmに拡大）[cite: 4]
    chk_header = Paragraph("<b>■ 工程チェックリスト</b>[cite: 4]", bold_style)
    story.append(chk_header)
    story.append(Spacer(1, 1 * mm))

    checklist_data = [
        [Paragraph("<b>【丁合】</b>[cite: 4]", bold_style), Paragraph("・背丁、ページの確認[cite: 4]<br/>・書名・台数・本数の確認[cite: 4]<br/>・作業終了後余り丁・表紙の整理[cite: 4]", chk_content_style), Paragraph("作業日:  年 月 日<br/>担当者印:[cite: 4]", chk_content_style)],
        [Paragraph("<b>【無線】</b>[cite: 4]", bold_style), Paragraph("・表紙(書名、巻数、天グル防止)の確認[cite: 4]<br/>・見返し・扉・口絵等の確認[cite: 4]<br/>・カット・アジロ・小口糊の確認[cite: 4]", chk_content_style), Paragraph("作業日:  年 月 日<br/>担当者印:[cite: 4]", chk_content_style)],
        [Paragraph("<b>【三方】</b>[cite: 4]", bold_style), Paragraph("・寸法・落としの確認[cite: 4]<br/>・アメ・カブリ・字切れ・トンボ残りの確認[cite: 4]<br/>・折込・汚れの確認[cite: 4]", chk_content_style), Paragraph("作業日:  年 月 日<br/>担当者印:[cite: 4]", chk_content_style)],
        [Paragraph("<b>【仕上げ】</b>[cite: 4]", bold_style), Paragraph("・背丁・ページの確認 / 付物等書名・ISBN・定価等の確認[cite: 4]<br/>・付物の印刷ヤレ確認 / 梱包の糊の具合・判子の状態の確認[cite: 4]<br/>・作業後の片付けの確認 / 製品の最終チェック[cite: 4]", chk_content_style), Paragraph("作業日:  年 月 日<br/>担当者印:[cite: 4]", chk_content_style)],
    ]

    chk_table = Table(checklist_data, colWidths=[22*mm, 112*mm, 60*mm], rowHeights=[17.5*mm, 17.5*mm, 17.5*mm, 17.5*mm])
    chk_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(chk_table)

    doc.build(story)
    with open(pdf_path, "rb") as f:
        return f.read()

# ==========================================
# 2. 面掛表・貼り込み表 PDF生成（高さ3倍広々版）
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
    center_style = ParagraphStyle('Center', fontName=DEFAULT_FONT, fontSize=7.5, leading=10, alignment=1)

    story = []

    # 1ページ目: 面掛表
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

    title_info_data = [
        [Paragraph("<b>書名</b>", bold_style), Paragraph(str(data.get('field-13', '')), normal_style), Paragraph("<b>面掛表</b><br/>受注番号: " + str(data.get('field-2', '')), title_style)],
        [Paragraph("<b>得意先 / 出版社</b>", bold_style), Paragraph(str(data.get('field-8', '')), normal_style), Paragraph(f"<b>ISBN:</b> {data.get('field-93', '')}", normal_style)],
        [Paragraph("<b>印刷所</b>", bold_style), Paragraph(str(data.get('field-14', '')), normal_style), Paragraph(f"<b>頁数:</b> {data.get('field-82', '')}頁", normal_style)]
    ]
    title_table = Table(title_info_data, colWidths=[30*mm, 94*mm, 70*mm])
    title_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(title_table)
    story.append(Spacer(1, 2 * mm))

    note_data = [[Paragraph(f"<b>注意事項:</b> {data.get('field-45', '特記事項なし')} | <b>背固仕様:</b> {data.get('field-43', 'アジロ無線')}", normal_style)]]
    note_table = Table(note_data, colWidths=[194*mm])
    note_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(note_table)
    story.append(Spacer(1, 3 * mm))

    # 面掛グリッド（高さ約22mm×4行で広々）[cite: 5]
    grid_rows = []
    for r in range(4):
        row_cells = []
        for c in range(10):
            num = r * 10 + c + 1
            row_cells.append(Paragraph(f"<b>{num}</b><br/><font size=6>折:   損:  </font>", center_style))
        grid_rows.append(row_cells)
    
    menkake_grid = Table(grid_rows, colWidths=[19.4*mm]*10, rowHeights=[22*mm]*4)
    menkake_grid.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('BACKGROUND', (0,0), (-1,-1), colors.white),
    ]))
    story.append(menkake_grid)
    
    story.append(PageBreak())

    # 2ページ目: 貼り込み表[cite: 6]
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

    pasting_headers = [Paragraph("<b>貼り込み箇所</b>", bold_style), Paragraph("<b>本数</b>", bold_style), Paragraph("<b>半端</b>", bold_style), Paragraph("<b>ヤレ枚数</b>", bold_style), Paragraph("<b>備考</b>", bold_style)]
    pasting_rows = [pasting_headers]
    
    for i in range(10):
        pasting_rows.append([
            Paragraph(f"折 〜 折", normal_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
        ])

    pasting_table = Table(pasting_rows, colWidths=[64*mm, 30*mm, 30*mm, 30*mm, 40*mm], rowHeights=[12*mm]*11)
    pasting_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.lightgrey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ALIGN', (1,1), (-1,-1), 'CENTER'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(pasting_table)

    doc.build(story)
    with open(pdf_path, "rb") as f:
        return f.read()

# ==========================================
# メイン画面処理 (URLパラメータ連動 ＆ ボタン選択)
# ==========================================
query_params = st.query_params
url_order_no = query_params.get("ord", None)

st.info("💡 @pocketの受注番号から、必要な帳票（製本進行カード または 面掛表）を選んでPDF出力できます。")

input_order_no = st.text_input("受注番号を入力", value=url_order_no if url_order_no else "")

if input_order_no:
    with st.spinner(f"受注番号 {input_order_no} のデータを@pocketから取得中..."):
        record_data = fetch_record_by_order(input_order_no)
        
    if record_data:
        st.success(f"✨ 対象レコードが見つかりました: **{record_data.get('field-13', '')}** ({record_data.get('field-8', '')})")
        
        st.markdown("---")
        st.subheader("🖨️ 出力する帳票を選択してください")
        
        col_btn1, col_btn2 = st.columns(2)
        
        # 1. 製本進行カード出力ボタン
        with col_btn1:
            st.markdown("#### 📋 製本進行カード")
            st.write("罫線高さを1.5倍に広げたチェックリスト付き")
            progress_pdf_bytes = generate_progress_card_pdf(record_data)
            st.download_button(
                label="📥 進行カードPDFをダウンロード",
                data=progress_pdf_bytes,
                file_name=f"製本進行カード_{input_order_no}.pdf",
                mime="application/pdf",
                type="primary",
                key="btn_progress"
            )
            
        # 2. 面掛表出力ボタン
        with col_btn2:
            st.markdown("#### 📋 面掛表・貼り込み表")
            st.write("高さを3倍に広げた面掛グリッド・貼り込み表")
            menkake_pdf_bytes = generate_menkake_pdf(record_data)
            st.download_button(
                label="📥 面掛表PDFをダウンロード",
                data=menkake_pdf_bytes,
                file_name=f"面掛表_{input_order_no}.pdf",
                mime="application/pdf",
                type="primary",
                key="btn_menkake"
            )
    else:
        st.error(f"❌ 受注番号「{input_order_no}」に一致するレコードが見つかりませんでした。")
else:
    st.warning("⚠️ 受注番号を入力してください（URLパラメータ `?ord=番号` からも自動読み込みされます）。")