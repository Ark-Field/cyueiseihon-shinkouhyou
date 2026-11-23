import io
import os
import requests
import pandas as pd
from datetime import datetime
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.pagebreak import Break
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
FONT_CHK = Font(name=EXCEL_FONT_NAME, size=8.5, bold=False, color="111827")

FILL_HEADER_GRAY = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
FILL_LABEL_BG = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
FILL_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

THIN_SIDE = Side(style="thin", color="6B7280")
MED_SIDE = Side(style="medium", color="111827")

BORDER_GRID = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

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
# 1-A. 製本進行カード Excel生成 (1シート内改ページ・全2頁構成)
# ==========================================
def generate_progress_card_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "製本進行カード"
    _configure_a4_portrait_print(ws, fit_h=2)

    # 18列幅設定
    ws.column_dimensions["A"].width = 4.25
    for idx in range(2, 19):
        ws.column_dimensions[get_column_letter(idx)].width = 13.0

    # --- 1頁目: ヘッダー (1行目) ---
    ws.row_dimensions[1].height = 24
    ws.merge_cells("A1:C1")
    ws.merge_cells("D1:F1")
    ws.merge_cells("G1:I1")
    ws.merge_cells("J1:O1")
    ws.merge_cells("P1:R1")

    ws["A1"] = f"納期: {_safe_str(data.get('field-52'))}"
    ws["D1"] = f"見本: {_safe_str(data.get('field-54'))}"
    ws["G1"] = f"配本: {_safe_str(data.get('field-56'))}"
    ws["J1"] = "製本進行カード (1/2)"
    ws["P1"] = f"受注番号: {_safe_str(data.get('field-2'))}"

    for c in range(1, 19):
        cell = ws.cell(row=1, column=c)
        cell.font = FONT_TITLE_CARD if 10 <= c <= 15 else FONT_BOLD
        cell.fill = FILL_HEADER_GRAY if 10 <= c <= 15 else FILL_LABEL_BG
        cell.alignment = ALIGN_CENTER if 10 <= c <= 15 else ALIGN_LEFT
    _apply_box_border(ws, 1, 1, 18, 1)

    # --- 2〜7行目: 基本・仕様情報 ---
    info_spec_rows = [
        (2, "書名", _safe_str(data.get("field-13")), "出版社", _safe_str(data.get("field-8"))),
        (3, "ISBN", _safe_str(data.get("field-93")), "印刷所", _safe_str(data.get("field-14"))),
        (
            4,
            "種別/刷数",
            f"{_safe_str(data.get('field-19'))} / {_safe_str(data.get('field-28'))}刷 (回:{_safe_str(data.get('field-29'))})",
            "定価",
            f"¥{_safe_str(data.get('field-20'))}",
        ),
        (
            5,
            "総頁数",
            f"{_safe_str(data.get('field-82'))} 頁",
            "印刷/製本数",
            f"印:{_safe_str(data.get('field-3'))}部 / 製:{_safe_str(data.get('field-4'))}部",
        ),
        (
            6,
            "仕上仕様",
            f"{_safe_str(data.get('field-41'))} / {_safe_str(data.get('field-42'))}",
            "背固仕様",
            _safe_str(data.get("field-43")),
        ),
        (
            7,
            "仕上寸法",
            f"天地:{_safe_str(data.get('field-24'))} / 左右:{_safe_str(data.get('field-25'))} ({_safe_str(data.get('field-44'))})",
            "付物/見返し",
            f"付:{_safe_str(data.get('field-49'))} / 見:{_safe_str(data.get('field-48'))}",
        ),
    ]

    for r, l_lbl, l_val, r_lbl, r_val in info_spec_rows:
        ws.row_dimensions[r].height = 20
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

    # --- 8行目: 口絵 / 折図、半製品預かり、予備(M8:R8) ---
    ws.row_dimensions[8].height = 20
    ws.merge_cells("A8:C8")
    ws.merge_cells("D8:I8")
    ws.merge_cells("J8:L8")
    ws.merge_cells("M8:R8")

    ws["A8"] = "口絵 / 折図"
    ws["D8"] = f"口絵:{_safe_str(data.get('field-65'))}(丁:{_safe_str(data.get('field-66'))}) / 折図:{_safe_str(data.get('field-69'))}"
    ws["J8"] = "半製品預かり"
    ws["M8"] = _safe_str(data.get("field-81"))

    for c in range(1, 19):
        cell = ws.cell(row=8, column=c)
        is_label = (1 <= c <= 3) or (10 <= c <= 12)
        cell.font = FONT_BOLD if is_label else FONT_NORMAL
        cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
        cell.alignment = ALIGN_LEFT
        cell.border = BORDER_GRID

    # --- 9行目: 袋払い (D9:I9 に 袋払い: {field-71} 表示) ＋ 予備 (J9:L9, M9:R9) ---
    ws.row_dimensions[9].height = 20
    ws.merge_cells("A9:C9")
    ws.merge_cells("D9:I9")
    ws.merge_cells("J9:L9")
    ws.merge_cells("M9:R9")

    ws["A9"] = "袋払い"
    ws["D9"] = f"袋払い: {_safe_str(data.get('field-71'))}"
    ws["J9"] = ""
    ws["M9"] = ""

    for c in range(1, 19):
        cell = ws.cell(row=9, column=c)
        is_label = (1 <= c <= 3) or (10 <= c <= 12)
        cell.font = FONT_BOLD if is_label else FONT_NORMAL
        cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
        cell.alignment = ALIGN_LEFT
        cell.border = BORDER_GRID

    # --- 10行目: 余白 ---
    ws.row_dimensions[10].height = 8

    # --- 11〜13行目: 数字マス目 (1〜54) ---
    num_ranges = [range(1, 19), range(19, 37), range(37, 55)]
    for idx, rng in enumerate(num_ranges):
        r = 11 + idx
        ws.row_dimensions[r].height = 34
        for col_offset, num in enumerate(rng):
            c = 1 + col_offset
            cell = ws.cell(row=r, column=c, value=num)
            cell.font = FONT_SMALL_BOLD
            cell.fill = FILL_WHITE
            cell.alignment = ALIGN_CENTER
            cell.border = BORDER_GRID

    # --- 14行目: 余白 ---
    ws.row_dimensions[14].height = 8

    # --- 15〜19行目: 手書き用エリア（丁合/無線/化粧/仕上/見本 ＋ 申送り） ---
    memo_labels = ["丁合", "無線", "化粧", "仕上", "見本"]
    for idx, lbl in enumerate(memo_labels):
        r = 15 + idx
        ws.row_dimensions[r].height = 38
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=18)

        ws.cell(row=r, column=1, value=lbl)
        ws.cell(row=r, column=4, value="申送り")
        ws.cell(row=r, column=5, value="")

        for c in range(1, 19):
            cell = ws.cell(row=r, column=c)
            if 1 <= c <= 3:
                cell.font = FONT_BOLD
                cell.fill = FILL_LABEL_BG
                cell.alignment = ALIGN_CENTER
            elif c == 4:
                cell.font = FONT_SMALL_BOLD
                cell.fill = FILL_LABEL_BG
                cell.alignment = ALIGN_CENTER
            else:
                cell.font = FONT_NORMAL
                cell.fill = FILL_WHITE
                cell.alignment = ALIGN_TOP_LEFT
            cell.border = BORDER_GRID

    # --- 20行目: 手書き用エリア（中永製本 ＋ フリー手書きスペース） ---
    ws.row_dimensions[20].height = 68
    ws.merge_cells("A20:C20")
    ws.merge_cells("D20:R20")
    ws["A20"] = "中永製本"
    ws["D20"] = ""

    for c in range(1, 19):
        cell = ws.cell(row=20, column=c)
        if 1 <= c <= 3:
            cell.font = FONT_BOLD
            cell.fill = FILL_LABEL_BG
            cell.alignment = ALIGN_CENTER
        else:
            cell.font = FONT_NORMAL
            cell.fill = FILL_WHITE
            cell.alignment = ALIGN_TOP_LEFT
        cell.border = BORDER_GRID

    # --- 1シート内改ページ設定（20行目の後に改ページを挿入し、21行目から2頁目） ---
    ws.row_breaks.append(Break(id=20))

    # ==========================================
    # 2頁目 (21行目〜): 特記事項 ＋ 工程チェックリスト
    # ==========================================
    ws.row_dimensions[21].height = 24
    ws.merge_cells("A21:O21")
    ws.merge_cells("P21:R21")
    ws["A21"] = f"製本進行カード (2/2) — 書名: {_safe_str(data.get('field-13'))}"
    ws["P21"] = f"受注番号: {_safe_str(data.get('field-2'))}"
    for c in range(1, 19):
        cell = ws.cell(row=21, column=c)
        cell.font = FONT_BOLD
        cell.fill = FILL_HEADER_GRAY if c <= 15 else FILL_LABEL_BG
        cell.alignment = ALIGN_LEFT
    _apply_box_border(ws, 1, 21, 18, 21)

    ws.row_dimensions[22].height = 8

    # 特記事項 (23〜24行目)
    ws.row_dimensions[23].height = 20
    ws.merge_cells("A23:R23")
    ws["A23"] = "特記事項 / 申し送り欄（フリー手書きスペース）"
    for c in range(1, 19):
        cell = ws.cell(row=23, column=c)
        cell.font = FONT_BOLD
        cell.fill = FILL_LABEL_BG
        cell.alignment = ALIGN_LEFT
        cell.border = BORDER_GRID

    ws.row_dimensions[24].height = 90
    ws.merge_cells("A24:R24")
    ws["A24"] = _safe_str(data.get("field-45"), "特記事項なし")
    for c in range(1, 19):
        cell = ws.cell(row=24, column=c)
        cell.font = FONT_NORMAL
        cell.fill = FILL_WHITE
        cell.alignment = ALIGN_TOP_LEFT
        cell.border = BORDER_GRID

    ws.row_dimensions[25].height = 10

    # 工程チェックリスト見出し (26行目)
    ws.row_dimensions[26].height = 20
    ws.merge_cells("A26:R26")
    ws["A26"] = "■ 工程チェックリスト"
    for c in range(1, 19):
        cell = ws.cell(row=26, column=c)
        cell.font = FONT_BOLD
        cell.alignment = ALIGN_LEFT

    # 工程チェックリスト (27〜30行目)
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
        r = 27 + idx
        ws.row_dimensions[r].height = 68
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
# 1-B. 製本進行カード PDF生成 (全2頁構成)
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
    normal_style = ParagraphStyle('CardNormal', fontName=DEFAULT_FONT, fontSize=8, leading=10.5)
    bold_style = ParagraphStyle('CardBold', fontName=DEFAULT_FONT, fontSize=8, leading=10.5)
    center_style = ParagraphStyle('CardCenter', fontName=DEFAULT_FONT, fontSize=7.5, leading=9, alignment=1)
    chk_content_style = ParagraphStyle('ChkContent', fontName=DEFAULT_FONT, fontSize=8, leading=13)

    story = []

    # --- 1頁目: ヘッダー ---
    header_data = [
        [
            Paragraph(f"<b>納期:</b> {_safe_str(data.get('field-52'))}", normal_style),
            Paragraph(f"<b>見本:</b> {_safe_str(data.get('field-54'))}", normal_style),
            Paragraph(f"<b>配本:</b> {_safe_str(data.get('field-56'))}", normal_style),
            Paragraph("製本進行カード (1/2)", title_style),
            Paragraph(f"<b>受注番号:</b> {_safe_str(data.get('field-2'))}", bold_style),
        ]
    ]
    header_table = Table(header_data, colWidths=[38*mm, 30*mm, 30*mm, 58*mm, 38*mm], rowHeights=[9.5*mm])
    header_table.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (3,0), (3,0), colors.lightgrey),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 2 * mm))

    # --- 1頁目: 基本・仕様・袋払い(field-71)・予備枠 ---
    info_spec_data = [
        [
            Paragraph("書名", bold_style),
            Paragraph(_safe_str(data.get('field-13')), normal_style),
            Paragraph("出版社", bold_style),
            Paragraph(_safe_str(data.get('field-8')), normal_style),
        ],
        [
            Paragraph("ISBN", bold_style),
            Paragraph(_safe_str(data.get('field-93')), normal_style),
            Paragraph("印刷所", bold_style),
            Paragraph(_safe_str(data.get('field-14')), normal_style),
        ],
        [
            Paragraph("種別/刷数", bold_style),
            Paragraph(f"{_safe_str(data.get('field-19'))} / {_safe_str(data.get('field-28'))}刷 (回:{_safe_str(data.get('field-29'))})", normal_style),
            Paragraph("定価", bold_style),
            Paragraph(f"¥{_safe_str(data.get('field-20'))}", normal_style),
        ],
        [
            Paragraph("総頁数", bold_style),
            Paragraph(f"{_safe_str(data.get('field-82'))} 頁", normal_style),
            Paragraph("印刷/製本数", bold_style),
            Paragraph(f"印:{_safe_str(data.get('field-3'))}部 / 製:{_safe_str(data.get('field-4'))}部", normal_style),
        ],
        [
            Paragraph("仕上仕様", bold_style),
            Paragraph(f"{_safe_str(data.get('field-41'))} / {_safe_str(data.get('field-42'))}", normal_style),
            Paragraph("背固仕様", bold_style),
            Paragraph(_safe_str(data.get('field-43')), normal_style),
        ],
        [
            Paragraph("仕上寸法", bold_style),
            Paragraph(f"天地:{_safe_str(data.get('field-24'))} / 左右:{_safe_str(data.get('field-25'))} ({_safe_str(data.get('field-44'))})", normal_style),
            Paragraph("付物/見返し", bold_style),
            Paragraph(f"付:{_safe_str(data.get('field-49'))} / 見:{_safe_str(data.get('field-48'))}", normal_style),
        ],
        [
            Paragraph("口絵 / 折図", bold_style),
            Paragraph(f"口絵:{_safe_str(data.get('field-65'))}(丁:{_safe_str(data.get('field-66'))}) / 折図:{_safe_str(data.get('field-69'))}", normal_style),
            Paragraph("半製品預かり", bold_style),
            Paragraph(_safe_str(data.get('field-81')), normal_style),
        ],
        [
            Paragraph("袋払い", bold_style),
            Paragraph(f"袋払い: {_safe_str(data.get('field-71'))}", normal_style),
            Paragraph("", bold_style),
            Paragraph("", normal_style),
        ],
    ]
    info_spec_table = Table(info_spec_data, colWidths=[24*mm, 70*mm, 24*mm, 76*mm], rowHeights=[7*mm] * 8)
    info_spec_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
        ('BACKGROUND', (2,0), (2,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(info_spec_table)
    story.append(Spacer(1, 2.5 * mm))

    # --- 1頁目: 数字マス目 (1〜54) ---
    row1_nums = [Paragraph(str(i), center_style) for i in range(1, 19)]
    row2_nums = [Paragraph(str(i), center_style) for i in range(19, 37)]
    row3_nums = [Paragraph(str(i), center_style) for i in range(37, 55)]

    num_grid_table = Table(
        [row1_nums, row2_nums, row3_nums],
        colWidths=[10.77 * mm] * 18,
        rowHeights=[12.5 * mm, 12.5 * mm, 12.5 * mm]
    )
    num_grid_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2),
        ('BACKGROUND', (0,0), (-1,-1), colors.white),
    ]))
    story.append(num_grid_table)
    story.append(Spacer(1, 2.5 * mm))

    # --- 1頁目: 手書き用エリア（丁合/無線/化粧/仕上/見本 ＋ 申送り、中永製本） ---
    handwriting_rows = [
        [Paragraph("<b>丁合</b>", bold_style), Paragraph("申送り", center_style), Paragraph("", normal_style)],
        [Paragraph("<b>無線</b>", bold_style), Paragraph("申送り", center_style), Paragraph("", normal_style)],
        [Paragraph("<b>化粧</b>", bold_style), Paragraph("申送り", center_style), Paragraph("", normal_style)],
        [Paragraph("<b>仕上</b>", bold_style), Paragraph("申送り", center_style), Paragraph("", normal_style)],
        [Paragraph("<b>見本</b>", bold_style), Paragraph("申送り", center_style), Paragraph("", normal_style)],
    ]
    handwriting_table = Table(handwriting_rows, colWidths=[22*mm, 18*mm, 154*mm], rowHeights=[15*mm] * 5)
    handwriting_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (1,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(handwriting_table)
    story.append(Spacer(1, 2 * mm))

    nakae_table = Table(
        [[Paragraph("<b>中永製本</b>", bold_style), Paragraph("", normal_style)]],
        colWidths=[22*mm, 172*mm],
        rowHeights=[24*mm]
    )
    nakae_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(nakae_table)

    # ==========================================
    # 2頁目: 特記事項 ＋ 工程チェックリスト
    # ==========================================
    story.append(PageBreak())

    page2_header = Table(
        [[
            Paragraph(f"<b>製本進行カード (2/2)</b> — {_safe_str(data.get('field-13'))}", bold_style),
            Paragraph(f"<b>受注番号:</b> {_safe_str(data.get('field-2'))}", bold_style),
        ]],
        colWidths=[150*mm, 44*mm],
        rowHeights=[9*mm]
    )
    page2_header.setStyle(TableStyle([
        ('BOX', (0,0), (-1,-1), 1, colors.black),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('BACKGROUND', (0,0), (0,0), colors.lightgrey),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(page2_header)
    story.append(Spacer(1, 3 * mm))

    note_data = [
        [Paragraph("特記事項 / 申し送り欄（フリー手書きスペース）", bold_style)],
        [Paragraph(_safe_str(data.get('field-45'), '特記事項なし') + "<br/><br/><br/><br/>", normal_style)]
    ]
    note_table = Table(note_data, colWidths=[194*mm], rowHeights=[8*mm, 34*mm])
    note_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 4),
    ]))
    story.append(note_table)
    story.append(Spacer(1, 3 * mm))

    story.append(Paragraph("<b>■ 工程チェックリスト</b>", bold_style))
    story.append(Spacer(1, 1.5 * mm))

    checklist_data = [
        [
            Paragraph("<b>【丁合】</b>", bold_style),
            Paragraph("・背丁、ページの確認<br/>・書名・台数・本数の確認<br/>・作業終了後余り丁・表紙の整理", chk_content_style),
            Paragraph("作業日:     年   月   日<br/>担当者印:", chk_content_style),
        ],
        [
            Paragraph("<b>【無線】</b>", bold_style),
            Paragraph("・表紙(書名、巻数、天グル防止)の確認<br/>・見返し・扉・口絵等の確認<br/>・カット・アジロ・小口糊の確認", chk_content_style),
            Paragraph("作業日:     年   月   日<br/>担当者印:", chk_content_style),
        ],
        [
            Paragraph("<b>【三方】</b>", bold_style),
            Paragraph("・寸法・落としの確認<br/>・アメ・カブリ・字切れ・トンボ残りの確認<br/>・折込・汚れの確認", chk_content_style),
            Paragraph("作業日:     年   月   日<br/>担当者印:", chk_content_style),
        ],
        [
            Paragraph("<b>【仕上げ】</b>", bold_style),
            Paragraph("・背丁・ページの確認 / 付物等書名・ISBN・定価等の確認<br/>・付物の印刷ヤレ確認 / 梱包の糊の具合・判子の状態の確認<br/>・作業後の片付けの確認 / 製品の最終チェック", chk_content_style),
            Paragraph("作業日:     年   月   日<br/>担当者印:", chk_content_style),
        ],
    ]

    chk_table = Table(checklist_data, colWidths=[24*mm, 114*mm, 56*mm], rowHeights=[34*mm, 34*mm, 34*mm, 38*mm])
    chk_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,-1), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 4),
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

        with col_btn1:
            st.markdown("#### 📋 製本進行カード（全2頁）")
            st.write("袋払い（field-71）・手書き欄・チェックリストを2頁に分割")
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

        with col_btn2:
            st.markdown("#### 📋 面掛表・貼り込み表")
            st.write("面掛表・折メーター表 ＆ 貼り込み表の2頁構成")
            menkake_pdf_bytes = generate_menkake_pdf(record_data)

            st.download_button(
                label="📥 面掛表・貼り込み表PDFをダウンロード",
                data=menkake_pdf_bytes,
                file_name=f"面掛表_貼り込み表_{input_order_no}.pdf",
                mime="application/pdf",
                type="primary",
                key="btn_menkake_pdf",
                use_container_width=True,
            )
    else:
        st.error(f"❌ 受注番号「{input_order_no}」に一致するレコードが見つかりませんでした。")
else:
    st.warning("⚠️ 受注番号を入力してください（URLパラメータ `?ord=番号` からも自動読み込みされます）。")