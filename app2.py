import io
import os
import requests
import pandas as pd
from datetime import datetime
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

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
# 日本語フォント登録 (PDF用)
# ==========================================
FONT_PATH = "NotoSansJP-Regular.ttf"
if os.path.exists(FONT_PATH):
    pdfmetrics.registerFont(TTFont('Japanese', FONT_PATH))
    DEFAULT_FONT = 'Japanese'
else:
    DEFAULT_FONT = 'Helvetica'

# ==========================================
# Excelスタイル定義 (openpyxl)
# ==========================================
EXCEL_FONT_NAME = "Meiryo UI"

FONT_TITLE_CARD = Font(name=EXCEL_FONT_NAME, size=13, bold=True, color="111827")
FONT_TITLE_SHEET = Font(name=EXCEL_FONT_NAME, size=15, bold=True, color="111827")
FONT_BOLD = Font(name=EXCEL_FONT_NAME, size=9, bold=True, color="111827")
FONT_NORMAL = Font(name=EXCEL_FONT_NAME, size=9, bold=False, color="111827")
FONT_SMALL_CENTER = Font(name=EXCEL_FONT_NAME, size=8, bold=False, color="111827")
FONT_SMALL_BOLD = Font(name=EXCEL_FONT_NAME, size=8, bold=True, color="111827")
FONT_CHK = Font(name=EXCEL_FONT_NAME, size=8, bold=False, color="111827")

FILL_HEADER_GRAY = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
FILL_LABEL_BG = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
FILL_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

THIN_SIDE = Side(style="thin", color="6B7280")
MED_SIDE = Side(style="medium", color="111827")

BORDER_GRID = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
BORDER_BOX = Border(left=MED_SIDE, right=MED_SIDE, top=MED_SIDE, bottom=MED_SIDE)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_TOP_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def _configure_a4_portrait_print(ws, fit_h=1):
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_h
    ws.page_margins.left = 0.25
    ws.page_margins.right = 0.25
    ws.page_margins.top = 0.25
    ws.page_margins.bottom = 0.25


def _apply_box_border(ws, min_col, min_row, max_col, max_row):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            top = MED_SIDE if r == min_row else THIN_SIDE
            bottom = MED_SIDE if r == max_row else THIN_SIDE
            left = MED_SIDE if c == min_col else THIN_SIDE
            right = MED_SIDE if c == max_col else THIN_SIDE
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)


def _safe_str(val, default=""):
    if val is None:
        return default
    return str(val)


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
# 1-A. 製本進行カード Excel生成 (openpyxl)
# ==========================================
def generate_progress_card_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "製本進行カード"
    _configure_a4_portrait_print(ws, fit_h=1)

    # 18列構成（数字マス目18列に合わせる）
    col_widths = [4.2] * 18
    for idx, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    # 1行目: ヘッダー (Cols 1-18)
    ws.row_dimensions[1].height = 24
    ws.merge_cells("A1:C1")
    ws.merge_cells("D1:F1")
    ws.merge_cells("G1:I1")
    ws.merge_cells("J1:O1")
    ws.merge_cells("P1:R1")

    ws["A1"] = f"納期: {_safe_str(data.get('field-52'))}"
    ws["D1"] = f"見本: {_safe_str(data.get('field-54'))}"
    ws["G1"] = f"配本: {_safe_str(data.get('field-56'))}"
    ws["J1"] = "製本進行カード"
    ws["P1"] = f"受注番号: {_safe_str(data.get('field-2'))}"

    for c in range(1, 19):
        cell = ws.cell(row=1, column=c)
        cell.font = FONT_TITLE_CARD if 10 <= c <= 15 else FONT_BOLD
        cell.fill = FILL_HEADER_GRAY if 10 <= c <= 15 else FILL_LABEL_BG
        cell.alignment = ALIGN_CENTER if 10 <= c <= 15 else ALIGN_LEFT
    _apply_box_border(ws, 1, 1, 18, 1)

    # 2〜5行目: 基本情報 (左ラベル A-C:1-3, 左値 D-I:4-9, 右ラベル J-L:10-12, 右値 M-R:13-18)
    info_rows = [
        ("書名", _safe_str(data.get("field-13")), "出版社", _safe_str(data.get("field-8"))),
        ("ISBN", _safe_str(data.get("field-93")), "印刷所", _safe_str(data.get("field-14"))),
        (
            "種別/刷数",
            f"{_safe_str(data.get('field-19'))} / {_safe_str(data.get('field-28'))}刷 (回:{_safe_str(data.get('field-29'))})",
            "定価",
            f"¥{_safe_str(data.get('field-20'))}",
        ),
        (
            "総頁数",
            f"{_safe_str(data.get('field-82'))} 頁",
            "印刷/製本数",
            f"印:{_safe_str(data.get('field-3'))}部 / 製:{_safe_str(data.get('field-4'))}部",
        ),
    ]

    for offset, (l_lbl, l_val, r_lbl, r_val) in enumerate(info_rows):
        r = 2 + offset
        ws.row_dimensions[r].height = 19
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=9)
        ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=12)
        ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=18)

        ws.cell(row=r, column=1, value=l_lbl)
        ws.cell(row=r, column=4, value=l_val)
        ws.cell(row=r, column=10, value=r_lbl)
        ws.cell(row=r, column=13, value=r_val)

        for c in range(1, 19):
            cell = ws.cell(row=r, column=c)
            is_label = (1 <= c <= 3) or (10 <= c <= 12)
            cell.font = FONT_BOLD if is_label else FONT_NORMAL
            cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
            cell.alignment = ALIGN_LEFT
            cell.border = BORDER_GRID

    # 6〜8行目: 仕様情報
    spec_rows = [
        (
            "仕上仕様",
            f"{_safe_str(data.get('field-41'))} / {_safe_str(data.get('field-42'))}",
            "背固仕様",
            _safe_str(data.get("field-43")),
        ),
        (
            "仕上寸法",
            f"天地:{_safe_str(data.get('field-24'))} / 左右:{_safe_str(data.get('field-25'))} ({_safe_str(data.get('field-44'))})",
            "付物/見返し",
            f"付:{_safe_str(data.get('field-49'))} / 見:{_safe_str(data.get('field-48'))}",
        ),
        (
            "口絵 / 折図",
            f"口絵:{_safe_str(data.get('field-65'))}(丁:{_safe_str(data.get('field-66'))}) / 折図:{_safe_str(data.get('field-69'))}",
            "半製品預かり",
            _safe_str(data.get("field-81")),
        ),
    ]

    for offset, (l_lbl, l_val, r_lbl, r_val) in enumerate(spec_rows):
        r = 6 + offset
        ws.row_dimensions[r].height = 19
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=9)
        ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=12)
        ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=18)

        ws.cell(row=r, column=1, value=l_lbl)
        ws.cell(row=r, column=4, value=l_val)
        ws.cell(row=r, column=10, value=r_lbl)
        ws.cell(row=r, column=13, value=r_val)

        for c in range(1, 19):
            cell = ws.cell(row=r, column=c)
            is_label = (1 <= c <= 3) or (10 <= c <= 12)
            cell.font = FONT_BOLD if is_label else FONT_NORMAL
            cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
            cell.alignment = ALIGN_LEFT
            cell.border = BORDER_GRID

    # 9行目: 余白
    ws.row_dimensions[9].height = 6

    # 10〜12行目: 数字マス目 (1〜54)
    num_ranges = [
        range(1, 19),
        range(19, 37),
        range(37, 55),
    ]
    for idx, rng in enumerate(num_ranges):
        r = 10 + idx
        ws.row_dimensions[r].height = 30
        for col_offset, num in enumerate(rng):
            c = 1 + col_offset
            cell = ws.cell(row=r, column=c, value=num)
            cell.font = FONT_SMALL_BOLD
            cell.fill = FILL_WHITE
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_GRID

    # 13行目: 余白
    ws.row_dimensions[13].height = 6

    # 14〜15行目: 特記事項
    ws.row_dimensions[14].height = 18
    ws.merge_cells("A14:R14")
    ws["A14"] = "特記事項 / 申し送り欄（フリー手書きスペース）"
    for c in range(1, 19):
        cell = ws.cell(row=14, column=c)
        cell.font = FONT_BOLD
        cell.fill = FILL_LABEL_BG
        cell.alignment = ALIGN_LEFT
        cell.border = BORDER_GRID

    ws.row_dimensions[15].height = 54
    ws.merge_cells("A15:R15")
    ws["A15"] = _safe_str(data.get("field-45"), "特記事項なし")
    for c in range(1, 19):
        cell = ws.cell(row=15, column=c)
        cell.font = FONT_NORMAL
        cell.fill = FILL_WHITE
        cell.alignment = ALIGN_TOP_LEFT
        cell.border = BORDER_GRID

    # 16行目: チェックリスト見出し
    ws.row_dimensions[16].height = 18
    ws.merge_cells("A16:R16")
    ws["A16"] = "■ 工程チェックリスト"
    for c in range(1, 19):
        cell = ws.cell(row=16, column=c)
        cell.font = FONT_BOLD
        cell.alignment = ALIGN_LEFT

    # 17〜20行目: 工程チェックリスト (工程名 A-C:1-3, 項目 D-M:4-13, 作業日/印 N-R:14-18)
    checklist_items = [
        (
            "【丁合】",
            "・背丁、ページの確認\n・書名・台数・本数の確認\n・作業終了後余り丁・表紙の整理",
            "作業日:     年   月   日\n担当者印:",
        ),
        (
            "【無線】",
            "・表紙(書名、巻数、天グル防止)の確認\n・見返し・扉・口絵等の確認\n・カット・アジロ・小口糊の確認",
            "作業日:     年   月   日\n担当者印:",
        ),
        (
            "【三方】",
            "・寸法・落としの確認\n・アメ・カブリ・字切れ・トンボ残りの確認\n・折込・汚れの確認",
            "作業日:     年   月   日\n担当者印:",
        ),
        (
            "【仕上げ】",
            "・背丁・ページの確認 / 付物等書名・ISBN・定価等の確認\n・付物の印刷ヤレ確認 / 梱包の糊の具合・判子の状態の確認\n・作業後の片付けの確認 / 製品の最終チェック",
            "作業日:     年   月   日\n担当者印:",
        ),
    ]

    for idx, (proc_name, proc_chk, proc_sign) in enumerate(checklist_items):
        r = 17 + idx
        ws.row_dimensions[r].height = 46
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=13)
        ws.merge_cells(start_row=r, start_column=14, end_row=r, end_column=18)

        ws.cell(row=r, column=1, value=proc_name)
        ws.cell(row=r, column=4, value=proc_chk)
        ws.cell(row=r, column=14, value=proc_sign)

        for c in range(1, 19):
            cell = ws.cell(row=r, column=c)
            if 1 <= c <= 3:
                cell.font = FONT_BOLD
                cell.fill = FILL_LABEL_BG
                cell.alignment = ALIGN_CENTER
            elif 4 <= c <= 13:
                cell.font = FONT_CHK
                cell.fill = FILL_WHITE
                cell.alignment = ALIGN_LEFT
            else:
                cell.font = FONT_CHK
                cell.fill = FILL_WHITE
                cell.alignment = ALIGN_TOP_LEFT
            cell.border = BORDER_GRID

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ==========================================
# 1-B. 製本進行カード PDF生成
# ==========================================
def generate_progress_card_pdf(data):
    pdf_path = "temp_progress_card.pdf"
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=8 * mm, leftMargin=8 * mm,
        topMargin=6 * mm, bottomMargin=6 * mm
    )

    title_style = ParagraphStyle('CardTitle', fontName=DEFAULT_FONT, fontSize=12, leading=15, alignment=1)
    normal_style = ParagraphStyle('CardNormal', fontName=DEFAULT_FONT, fontSize=7.5, leading=9.5)
    bold_style = ParagraphStyle('CardBold', fontName=DEFAULT_FONT, fontSize=7.5, leading=9.5)
    center_style = ParagraphStyle('CardCenter', fontName=DEFAULT_FONT, fontSize=7, leading=8, alignment=1)
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

    # 基本情報
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
        ('BACKGROUND', (0,0), (-1,-1), colors.whitesmoke),
        ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(info_table)
    story.append(Spacer(1, 1.5 * mm))

    # 仕様情報
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
        [Paragraph(str(data.get('field-45', '特記事項なし')) + "<br/><br/><br/>", normal_style)]
    ]
    note_table = Table(note_data, colWidths=[194*mm])
    note_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(note_table)
    story.append(Spacer(1, 1.5 * mm))

    # 工程チェックリスト
    chk_header = Paragraph("<b>■ 工程チェックリスト</b>", bold_style)
    story.append(chk_header)
    story.append(Spacer(1, 1 * mm))

    checklist_data = [
        [Paragraph("<b>【丁合】</b>", bold_style), Paragraph("・背丁、ページの確認<br/>・書名・台数・本数の確認<br/>・作業終了後余り丁・表紙の整理", chk_content_style), Paragraph("作業日:  年 月 日<br/>担当者印:", chk_content_style)],
        [Paragraph("<b>【無線】</b>", bold_style), Paragraph("・表紙(書名、巻数、天グル防止)の確認<br/>・見返し・扉・口絵等の確認<br/>・カット・アジロ・小口糊の確認", chk_content_style), Paragraph("作業日:  年 月 日<br/>担当者印:", chk_content_style)],
        [Paragraph("<b>【三方】</b>", bold_style), Paragraph("・寸法・落としの確認<br/>・アメ・カブリ・字切れ・トンボ残りの確認<br/>・折込・汚れの確認", chk_content_style), Paragraph("作業日:  年 月 日<br/>担当者印:", chk_content_style)],
        [Paragraph("<b>【仕上げ】</b>", bold_style), Paragraph("・背丁・ページの確認 / 付物等書名・ISBN・定価等の確認<br/>・付物の印刷ヤレ確認 / 梱包の糊の具合・判子の状態の確認<br/>・作業後の片付けの確認 / 製品の最終チェック", chk_content_style), Paragraph("作業日:  年 月 日<br/>担当者印:", chk_content_style)],
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
# 2-A. 面掛表・貼り込み表 Excel生成 (openpyxl)
# ==========================================
def generate_menkake_excel(data):
    wb = Workbook()

    # ------------------------------------------
    # シート1: 面掛表 ＆ 折メーター表 (16列構成)
    # ------------------------------------------
    ws1 = wb.active
    ws1.title = "面掛表"
    _configure_a4_portrait_print(ws1, fit_h=1)

    col_widths_1 = [
        4.5, 8.0, 11.0, 5.5,
        4.5, 8.0, 11.0, 5.5,
        4.5, 8.0, 11.0, 5.5,
        4.5, 8.0, 11.0, 5.5,
    ]
    for idx, w in enumerate(col_widths_1, start=1):
        ws1.column_dimensions[get_column_letter(idx)].width = w

    # タイトル
    ws1.row_dimensions[1].height = 24
    ws1.merge_cells("A1:P1")
    ws1["A1"] = "面掛表"
    ws1["A1"].font = FONT_TITLE_SHEET
    ws1["A1"].alignment = ALIGN_CENTER

    # 受注番号 ＆ ISBN
    ws1.row_dimensions[2].height = 18
    ws1.merge_cells("A2:H2")
    ws1.merge_cells("I2:P2")
    ws1["A2"] = f"受注番号: {_safe_str(data.get('field-2'))}"
    ws1["I2"] = f"ISBN: {_safe_str(data.get('field-93'))}"
    for c in range(1, 17):
        cell = ws1.cell(row=2, column=c)
        cell.font = FONT_BOLD
        cell.alignment = ALIGN_LEFT
    _apply_box_border(ws1, 1, 2, 16, 2)

    # 書名（手書きスペース確保）
    ws1.row_dimensions[3].height = 18
    ws1.merge_cells("A3:P3")
    ws1["A3"] = "書名:"
    for c in range(1, 17):
        cell = ws1.cell(row=3, column=c)
        cell.font = FONT_BOLD
        cell.alignment = ALIGN_LEFT
    _apply_box_border(ws1, 1, 3, 16, 3)

    # 面掛テーブル (4行目ヘッダー + 5〜14行の10行)
    # A-D(1-4):面掛1, E-F(5-6):摘要1, G-J(7-10):面掛2, K-L(11-12):摘要2, M-P(13-16):備考
    ws1.row_dimensions[4].height = 18
    ws1.merge_cells("A4:D4")
    ws1.merge_cells("E4:F4")
    ws1.merge_cells("G4:J4")
    ws1.merge_cells("K4:L4")
    ws1.merge_cells("M4:P4")
    ws1["A4"] = "面掛"
    ws1["E4"] = "摘要"
    ws1["G4"] = "面掛"
    ws1["K4"] = "摘要"
    ws1["M4"] = "備考"

    for c in range(1, 17):
        cell = ws1.cell(row=4, column=c)
        cell.font = FONT_BOLD
        cell.fill = FILL_LABEL_BG if c <= 4 else FILL_WHITE
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_GRID

    for i in range(10):
        r = 5 + i
        ws1.row_dimensions[r].height = 18
        ws1.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws1.merge_cells(start_row=r, start_column=5, end_row=r, end_column=6)
        ws1.merge_cells(start_row=r, start_column=7, end_row=r, end_column=10)
        ws1.merge_cells(start_row=r, start_column=11, end_row=r, end_column=12)
        ws1.merge_cells(start_row=r, start_column=13, end_row=r, end_column=16)

        ws1.cell(row=r, column=1, value="  ×  ")
        ws1.cell(row=r, column=7, value="  ×  ")

        for c in range(1, 17):
            cell = ws1.cell(row=r, column=c)
            cell.font = FONT_NORMAL
            cell.alignment = ALIGN_CENTER if c in (1, 7) else ALIGN_LEFT
            cell.border = BORDER_GRID

    # 注意事項1 (15行目)
    ws1.row_dimensions[15].height = 54
    ws1.merge_cells("A15:P15")
    ws1["A15"] = "注意事項:"
    for c in range(1, 17):
        cell = ws1.cell(row=15, column=c)
        cell.font = FONT_BOLD
        cell.alignment = ALIGN_TOP_LEFT
    _apply_box_border(ws1, 1, 15, 16, 15)

    # 16行目: 余白
    ws1.row_dimensions[16].height = 10

    # 17行目: 作業日
    ws1.row_dimensions[17].height = 18
    ws1.merge_cells("A17:P17")
    ws1["A17"] = "作業日:      年      月      日"
    for c in range(1, 17):
        cell = ws1.cell(row=17, column=c)
        cell.font = FONT_BOLD
        cell.alignment = ALIGN_LEFT
    _apply_box_border(ws1, 1, 17, 16, 17)

    # 18〜28行目: 折メーター表 (16列 × 11行)
    ws1.row_dimensions[18].height = 17
    for block in range(4):
        base_c = block * 4 + 1
        ws1.cell(row=18, column=base_c, value="No").font = FONT_SMALL_BOLD
        ws1.cell(row=18, column=base_c + 1, value="担当").font = FONT_SMALL_BOLD
        ws1.cell(row=18, column=base_c + 2, value="折メーター").font = FONT_SMALL_BOLD
        ws1.cell(row=18, column=base_c + 3, value="折損").font = FONT_SMALL_BOLD

    for c in range(1, 17):
        cell = ws1.cell(row=18, column=c)
        cell.fill = FILL_LABEL_BG if c in (1, 5, 9, 13) else FILL_WHITE
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_GRID

    for r_idx in range(10):
        r = 19 + r_idx
        ws1.row_dimensions[r].height = 16
        for block in range(4):
            base_c = block * 4 + 1
            no_val = r_idx + 1 + block * 10
            c_no = ws1.cell(row=r, column=base_c, value=no_val)
            c_no.font = FONT_SMALL_CENTER
            c_no.alignment = ALIGN_CENTER
            c_no.border = BORDER_GRID

            for sub_c in range(1, 4):
                cell = ws1.cell(row=r, column=base_c + sub_c, value="")
                cell.font = FONT_SMALL_CENTER
                cell.alignment = ALIGN_CENTER
                cell.border = BORDER_GRID

    # 注意事項2 (29行目)
    ws1.row_dimensions[29].height = 48
    ws1.merge_cells("A29:P29")
    ws1["A29"] = "注意事項:"
    for c in range(1, 17):
        cell = ws1.cell(row=29, column=c)
        cell.font = FONT_BOLD
        cell.alignment = ALIGN_TOP_LEFT
    _apply_box_border(ws1, 1, 29, 16, 29)

    # ------------------------------------------
    # シート2: 貼り込み表
    # ------------------------------------------
    ws2 = wb.create_sheet(title="貼り込み表")
    _configure_a4_portrait_print(ws2, fit_h=1)

    col_widths_2 = [24.0, 10.0, 10.0, 10.0, 24.0]
    for idx, w in enumerate(col_widths_2, start=1):
        ws2.column_dimensions[get_column_letter(idx)].width = w

    ws2.row_dimensions[1].height = 24
    ws2.merge_cells("A1:E1")
    ws2["A1"] = "貼り込み表"
    ws2["A1"].font = FONT_TITLE_SHEET
    ws2["A1"].alignment = ALIGN_CENTER

    ws2.row_dimensions[2].height = 20
    ws2.merge_cells("A2:B2")
    ws2["A2"] = "書名:"
    ws2["C2"] = f"部数: {_safe_str(data.get('field-4', '500'))} 部"
    ws2.merge_cells("D2:E2")
    ws2["D2"] = "作業日:    年    月    日"

    for c in range(1, 6):
        cell = ws2.cell(row=2, column=c)
        cell.font = FONT_BOLD
        cell.fill = FILL_LABEL_BG
        cell.alignment = ALIGN_LEFT
        cell.border = BORDER_GRID

    headers_2 = ["貼り込み箇所", "本数", "半端", "ヤレ枚数", "備考"]
    ws2.row_dimensions[3].height = 19
    for c, h_text in enumerate(headers_2, start=1):
        cell = ws2.cell(row=3, column=c, value=h_text)
        cell.font = FONT_BOLD
        cell.fill = FILL_LABEL_BG if c == 1 else FILL_WHITE
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_GRID

    for i in range(24):
        r = 4 + i
        ws2.row_dimensions[r].height = 21
        c1 = ws2.cell(row=r, column=1, value="    折〜    折")
        c1.font = FONT_NORMAL
        c1.alignment = ALIGN_LEFT
        c1.border = BORDER_GRID

        for c, unit in [(2, "本"), (3, "本"), (4, "枚")]:
            cell = ws2.cell(row=r, column=c, value=unit)
            cell.font = FONT_NORMAL
            cell.alignment = ALIGN_RIGHT
            cell.border = BORDER_GRID

        c5 = ws2.cell(row=r, column=5, value="")
        c5.font = FONT_NORMAL
        c5.alignment = ALIGN_LEFT
        c5.border = BORDER_GRID

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ==========================================
# 2-B. 面掛表・貼り込み表 PDF生成（2ページ構成）
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

    # ==========================================
    # 【1ページ目】面掛表 ＆ 表2
    # ==========================================
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

    # ==========================================
    # 下半部分：表2（16列構成）
    # ==========================================
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

    sub_col_widths = [
        7.5*mm, 13*mm, 19*mm, 9*mm,
        7.5*mm, 13*mm, 19*mm, 9*mm,
        7.5*mm, 13*mm, 19*mm, 9*mm,
        7.5*mm, 13*mm, 19*mm, 9*mm
    ]
    sub_table = Table(
        sub_rows,
        colWidths=sub_col_widths,
        rowHeights=[5.2*mm] * 11
    )
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

    # ==========================================
    # 【2ページ目】貼り込み表
    # ==========================================
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
# メイン画面処理 (URLパラメータ連動 ＆ ボタン選択)
# ==========================================
query_params = st.query_params
url_order_no = query_params.get("ord", None)

st.info("💡 @pocketの受注番号から、必要な帳票（製本進行カード または 面掛表・貼り込み表）を選んでPDF・Excel出力できます。")

input_order_no = st.text_input("受注番号を入力", value=url_order_no if url_order_no else "")

if input_order_no:
    with st.spinner(f"受注番号 {input_order_no} のデータを@pocketから取得中..."):
        record_data = fetch_record_by_order(input_order_no)

    if record_data:
        st.success(f"✨ 対象レコードが見つかりました: **{record_data.get('field-13', '')}** ({record_data.get('field-8', '')})")

        st.markdown("---")
        st.subheader("🖨️ 出力する帳票を選択してください")

        col_btn1, col_btn2 = st.columns(2)

        # 1. 製本進行カード出力（PDF + Excel）
        with col_btn1:
            st.markdown("#### 📋 製本進行カード")
            st.write("数字マス目（1〜54）と工程チェックリスト付き")
            progress_pdf_bytes = generate_progress_card_pdf(record_data)
            progress_xlsx_bytes = generate_progress_card_excel(record_data)

            st.download_button(
                label="📥 進行カードPDFをダウンロード",
                data=progress_pdf_bytes,
                file_name=f"製本進行カード_{input_order_no}.pdf",
                mime="application/pdf",
                type="primary",
                key="btn_progress_pdf",
                use_container_width=True,
            )
            st.download_button(
                label="📊 進行カードExcelをダウンロード",
                data=progress_xlsx_bytes,
                file_name=f"製本進行カード_{input_order_no}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_progress_xlsx",
                use_container_width=True,
            )

        # 2. 面掛表・貼り込み表出力（PDF + Excel）
        with col_btn2:
            st.markdown("#### 📋 面掛表・貼り込み表")
            st.write("面掛表・折メーター表 ＆ 貼り込み表の2構成")
            menkake_pdf_bytes = generate_menkake_pdf(record_data)
            menkake_xlsx_bytes = generate_menkake_excel(record_data)

            st.download_button(
                label="📥 面掛表・貼り込み表PDFをダウンロード",
                data=menkake_pdf_bytes,
                file_name=f"面掛表_貼り込み表_{input_order_no}.pdf",
                mime="application/pdf",
                type="primary",
                key="btn_menkake_pdf",
                use_container_width=True,
            )
            st.download_button(
                label="📊 面掛表・貼り込み表Excelをダウンロード",
                data=menkake_xlsx_bytes,
                file_name=f"面掛表_貼り込み表_{input_order_no}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="btn_menkake_xlsx",
                use_container_width=True,
            )
    else:
        st.error(f"❌ 受注番号「{input_order_no}」に一致するレコードが見つかりませんでした。")
else:
    st.warning("⚠️ 受注番号を入力してください（URLパラメータ `?ord=番号` からも自動読み込みされます）。")