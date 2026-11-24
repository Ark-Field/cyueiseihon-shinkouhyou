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
# 共通データ定義（総頁数折数表記・2頁目チェックリスト）
# ==========================================
FOLD_PAGE_GUIDE = "32p/ 16P／ 8P／ 4P／ 2P／"

CHECKLIST_SECTIONS = [
    (
        "【丁合】",
        [
            "・背丁、ページの確認",
            "・書名、台数、本数 の確認",
            "・作業終了後 余り丁 表紙の整理",
        ],
    ),
    (
        "【無線】",
        [
            "・表紙(書名、巻数、天グル防止)の確認",
            "・見返し・扉・口絵等の確認",
            "・カット・アジロ・小口糊の確認",
        ],
    ),
    (
        "【三方】",
        [
            "・寸法・落としの確認",
            "・アメ・カブリ・字切れ・トンボ残りの確認",
            "・折込・汚れの確認",
        ],
    ),
    (
        "【仕上げ】",
        [
            "・背丁・ページの確認 /",
            "・付物等書名・ISBN・定価等の確認",
            "・付物の印刷ヤレ確認",
            "・梱包の糊の具合・判子の状態の確認",
            "・作業後の片付けの確認",
            "・製品の最終チェック",
        ],
    ),
]


def _safe_str(val, default=""):
    if val is None:
        return default
    return str(val)


def _format_total_pages(data):
    raw_pages = _safe_str(data.get("field-82")).strip()
    page_part = f"{raw_pages} 頁 " if raw_pages else ""
    return f"{page_part}{FOLD_PAGE_GUIDE}"


def _format_fold_breakdown(data):
    parts = [
        f"32P:{_safe_str(data.get('field-30'), '-')}",
        f"16P:{_safe_str(data.get('field-31'), '-')}",
        f"8P:{_safe_str(data.get('field-32'), '-')}",
        f"4P:{_safe_str(data.get('field-33'), '-')}",
        f"2P:{_safe_str(data.get('field-34'), '-')}",
    ]
    return " / ".join(parts)


# ==========================================
# Excelスタイル定義 (openpyxl)
# ==========================================
EXCEL_FONT_NAME = "Meiryo UI"

FONT_TITLE_CARD = Font(name=EXCEL_FONT_NAME, size=13, bold=True, color="111827")
FONT_TITLE_MENKAKE = Font(name=EXCEL_FONT_NAME, size=13, bold=True, color="111827")
FONT_BOLD = Font(name=EXCEL_FONT_NAME, size=9, bold=True, color="111827")
FONT_NORMAL = Font(name=EXCEL_FONT_NAME, size=9, bold=False, color="111827")
FONT_SMALL_BOLD = Font(name=EXCEL_FONT_NAME, size=8, bold=True, color="111827")
FONT_SMALL = Font(name=EXCEL_FONT_NAME, size=8, bold=False, color="111827")
FONT_CHK = Font(name=EXCEL_FONT_NAME, size=8.5, bold=False, color="111827")

FILL_HEADER_GRAY = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
FILL_LABEL_BG = PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
FILL_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")

THIN_SIDE = Side(style="thin", color="6B7280")
MED_SIDE = Side(style="medium", color="111827")
DOUBLE_SIDE = Side(style="double", color="111827")

BORDER_GRID = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)
BORDER_CHK_HEADER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=DOUBLE_SIDE)

ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_TOP_LEFT = Alignment(horizontal="left", vertical="top", wrap_text=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center", wrap_text=True)


def _configure_a4_portrait_print(ws, fit_h=2, margin_lr=0.2, margin_tb=0.2):
    ws.views.sheetView[0].showGridLines = True
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = fit_h
    ws.page_margins.left = margin_lr
    ws.page_margins.right = margin_lr
    ws.page_margins.top = margin_tb
    ws.page_margins.bottom = margin_tb


def _apply_box_border(ws, min_col, min_row, max_col, max_row):
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            top = MED_SIDE if r == min_row else THIN_SIDE
            bottom = MED_SIDE if r == max_row else THIN_SIDE
            left = MED_SIDE if c == min_col else THIN_SIDE
            right = MED_SIDE if c == max_col else THIN_SIDE
            cell.border = Border(top=top, bottom=bottom, left=left, right=right)


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
    _configure_a4_portrait_print(ws, fit_h=2, margin_lr=0.25, margin_tb=0.25)

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
            _format_total_pages(data),
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

    # --- 8行目: 口絵 / 折図、半製品預かり ---
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

    # --- 9行目: 袋払い (field-71) ＋ 予備枠 ---
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

    # --- 20行目: 手書き用エリア（中永製本） ---
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

    # 1シート内改ページ（20行目の後）
    ws.row_breaks.append(Break(id=20))

    # ==========================================
    # 2頁目 (21行目〜): 特記事項 ＋ 工程チェックリスト
    # ==========================================
    ws.row_dimensions[21].height = 20
    ws.merge_cells("A21:R21")
    ws["A21"] = "特記事項 / 申し送り欄（フリー手書きスペース）"
    for c in range(1, 19):
        cell = ws.cell(row=21, column=c)
        cell.font = FONT_BOLD
        cell.fill = FILL_LABEL_BG
        cell.alignment = ALIGN_LEFT
        cell.border = BORDER_GRID

    ws.row_dimensions[22].height = 70
    ws.merge_cells("A22:R22")
    ws["A22"] = _safe_str(data.get("field-45"), "特記事項なし")
    for c in range(1, 19):
        cell = ws.cell(row=22, column=c)
        cell.font = FONT_NORMAL
        cell.fill = FILL_WHITE
        cell.alignment = ALIGN_TOP_LEFT
        cell.border = BORDER_GRID

    ws.row_dimensions[23].height = 8

    curr_r = 24
    for sec_title, items in CHECKLIST_SECTIONS:
        start_r = curr_r
        end_r = curr_r + len(items)

        ws.merge_cells(start_row=start_r, start_column=1, end_row=end_r, end_column=3)
        ws.cell(row=start_r, column=1, value=sec_title)

        ws.row_dimensions[start_r].height = 28
        ws.merge_cells(start_row=start_r, start_column=4, end_row=start_r, end_column=18)
        ws.cell(row=start_r, column=4, value="作業日：  年  月  日\n担当者印：")

        for c in range(1, 19):
            cell = ws.cell(row=start_r, column=c)
            cell.font = FONT_BOLD if 1 <= c <= 3 else FONT_CHK
            cell.fill = FILL_LABEL_BG if 1 <= c <= 3 else FILL_WHITE
            cell.alignment = ALIGN_CENTER if 1 <= c <= 3 else ALIGN_LEFT
            cell.border = BORDER_CHK_HEADER

        for idx, item_text in enumerate(items):
            r = start_r + 1 + idx
            ws.row_dimensions[r].height = 24
            ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=13)
            ws.merge_cells(start_row=r, start_column=14, end_row=r, end_column=18)

            ws.cell(row=r, column=4, value=item_text)
            ws.cell(row=r, column=14, value="")

            for c in range(1, 19):
                cell = ws.cell(row=r, column=c)
                if 1 <= c <= 3:
                    cell.font = FONT_BOLD
                    cell.fill = FILL_LABEL_BG
                    cell.alignment = ALIGN_CENTER
                else:
                    cell.font = FONT_CHK
                    cell.fill = FILL_WHITE
                    cell.alignment = ALIGN_LEFT
                cell.border = BORDER_GRID

        curr_r = end_r + 1

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
    chk_style = ParagraphStyle('ChkItem', fontName=DEFAULT_FONT, fontSize=8, leading=10.5)

    story = []

    # --- 1頁目 ---
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
            Paragraph(_format_total_pages(data), normal_style),
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

    # --- 2頁目: 特記事項 ＋ 工程チェックリスト ---
    story.append(PageBreak())

    note_data = [
        [Paragraph("特記事項 / 申し送り欄（フリー手書きスペース）", bold_style)],
        [Paragraph(_safe_str(data.get('field-45'), '特記事項なし') + "<br/><br/><br/>", normal_style)]
    ]
    note_table = Table(note_data, colWidths=[194*mm], rowHeights=[7*mm, 24*mm])
    note_table.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('BACKGROUND', (0,0), (0,0), colors.whitesmoke),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('PADDING', (0,0), (-1,-1), 3),
    ]))
    story.append(note_table)
    story.append(Spacer(1, 2 * mm))

    chk_rows = []
    chk_style_cmds = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('PADDING', (0,0), (-1,-1), 2.5),
    ]
    row_heights = []
    cur_idx = 0

    for sec_title, items in CHECKLIST_SECTIONS:
        start_idx = cur_idx
        end_idx = cur_idx + len(items)

        chk_rows.append([
            Paragraph(f"<b>{sec_title}</b>", bold_style),
            Paragraph("作業日：  年  月  日<br/>担当者印：", chk_style),
            Paragraph("", chk_style),
        ])
        row_heights.append(9.5 * mm)
        chk_style_cmds.append(('BACKGROUND', (0, cur_idx), (0, cur_idx), colors.whitesmoke))
        chk_style_cmds.append(('ALIGN', (0, cur_idx), (0, cur_idx), 'CENTER'))
        chk_style_cmds.append(('LINEBELOW', (1, cur_idx), (2, cur_idx), 1.2, colors.black))
        cur_idx += 1

        for item_text in items:
            chk_rows.append([
                Paragraph("", bold_style),
                Paragraph(item_text, chk_style),
                Paragraph("", chk_style),
            ])
            row_heights.append(8.2 * mm)
            chk_style_cmds.append(('BACKGROUND', (0, cur_idx), (0, cur_idx), colors.whitesmoke))
            chk_style_cmds.append(('ALIGN', (0, cur_idx), (0, cur_idx), 'CENTER'))
            cur_idx += 1

        chk_style_cmds.append(('SPAN', (0, start_idx), (0, end_idx)))

    chk_table = Table(chk_rows, colWidths=[24*mm, 118*mm, 52*mm], rowHeights=row_heights)
    chk_table.setStyle(TableStyle(chk_style_cmds))
    story.append(chk_table)

    doc.build(story)
    with open(pdf_path, "rb") as f:
        return f.read()


# ==========================================
# 2-A. 面掛表・貼り込み表 Excel生成 (1シート内改ページ・全2頁構成)
# ==========================================
def generate_menkake_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.title = "面掛表_貼り込み表"
    _configure_a4_portrait_print(ws, fit_h=2, margin_lr=0.2, margin_tb=0.18)

    # 16列共通グリッド（下段4ブロック×4列=16列を基準）
    # [No, 担当, 折メーター, 折損] × 4
    block_widths = [3.5, 6.2, 8.8, 4.3]
    for block_idx in range(4):
        for offset, w in enumerate(block_widths):
            col_idx = block_idx * 4 + offset + 1
            ws.column_dimensions[get_column_letter(col_idx)].width = w

    # ==========================================
    # 1頁目: 面掛表 ＆ 表組①(0.9cm) ＆ 注意事項(2cm) ＆ 作業日 ＆ 表組②(1.2cm) ＆ 注意事項(2cm)
    # ==========================================
    # --- 1行目: タイトル ---
    ws.row_dimensions[1].height = 20
    ws.merge_cells("A1:P1")
    ws["A1"] = "面 掛 表"
    for c in range(1, 17):
        cell = ws.cell(row=1, column=c)
        cell.font = FONT_TITLE_MENKAKE
        cell.fill = FILL_HEADER_GRAY
        cell.alignment = ALIGN_CENTER
    _apply_box_border(ws, 1, 1, 16, 1)

    # --- 2行目: 受注番号 / 出版社名 ---
    ws.row_dimensions[2].height = 16.5
    ws.merge_cells("A2:B2")
    ws.merge_cells("C2:H2")
    ws.merge_cells("I2:J2")
    ws.merge_cells("K2:P2")

    ws["A2"] = "受注番号"
    ws["C2"] = _safe_str(data.get("field-2"))
    ws["I2"] = "出版社名"
    ws["K2"] = _safe_str(data.get("field-8"))

    for c in range(1, 17):
        cell = ws.cell(row=2, column=c)
        is_label = c in (1, 2, 9, 10)
        cell.font = FONT_SMALL_BOLD if is_label else FONT_SMALL
        cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
        cell.alignment = ALIGN_CENTER if is_label else ALIGN_LEFT
        cell.border = BORDER_GRID

    # --- 3行目: 書名（field-13 自動印字） ---
    ws.row_dimensions[3].height = 17
    ws.merge_cells("A3:B3")
    ws.merge_cells("C3:P3")
    ws["A3"] = "書名"
    ws["C3"] = _safe_str(data.get("field-13"))
    for c in range(1, 17):
        cell = ws.cell(row=3, column=c)
        is_label = c in (1, 2)
        cell.font = FONT_SMALL_BOLD if is_label else FONT_BOLD
        cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
        cell.alignment = ALIGN_CENTER if is_label else ALIGN_LEFT
        cell.border = BORDER_GRID

    # --- 4行目: 総頁数 / 折数（32P,16P,8P,4P,2P）/ 合計台数 ---
    ws.row_dimensions[4].height = 17
    ws.merge_cells("A4:B4")
    ws.merge_cells("C4:D4")
    ws.merge_cells("E4:F4")
    ws.merge_cells("G4:M4")
    ws.merge_cells("N4:O4")
    ws["A4"] = "総頁数"
    ws["C4"] = _safe_str(data.get("field-82"))
    ws["E4"] = "折数"
    ws["G4"] = _format_fold_breakdown(data)
    ws["N4"] = "合計台数"
    ws["P4"] = _safe_str(data.get("field-35"))

    for c in range(1, 17):
        cell = ws.cell(row=4, column=c)
        is_label = c in (1, 2, 5, 6, 14, 15)
        cell.font = FONT_SMALL_BOLD if is_label else FONT_SMALL
        cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
        if is_label:
            cell.alignment = ALIGN_CENTER
        elif c in (3, 4, 16):
            cell.alignment = ALIGN_CENTER
        else:
            cell.alignment = ALIGN_LEFT
        cell.border = BORDER_GRID

    # --- 5行目: 表組① ヘッダー（面掛｜備考｜面掛｜備考） ---
    ws.row_dimensions[5].height = 16.5
    ws.merge_cells("A5:D5")
    ws.merge_cells("E5:H5")
    ws.merge_cells("I5:L5")
    ws.merge_cells("M5:P5")

    ws["A5"] = "面掛"
    ws["E5"] = "備考"
    ws["I5"] = "面掛"
    ws["M5"] = "備考"

    for c in range(1, 17):
        cell = ws.cell(row=5, column=c)
        cell.font = FONT_SMALL_BOLD
        cell.fill = FILL_LABEL_BG
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_GRID

    # --- 6〜15行目: 表組①（手書き用・枠のみ・4列×10段・行高さ 0.9cm = 25.5pt） ---
    for i in range(10):
        r = 6 + i
        ws.row_dimensions[r].height = 25.5
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
        ws.merge_cells(start_row=r, start_column=5, end_row=r, end_column=8)
        ws.merge_cells(start_row=r, start_column=9, end_row=r, end_column=12)
        ws.merge_cells(start_row=r, start_column=13, end_row=r, end_column=16)

        for start_c in (1, 5, 9, 13):
            ws.cell(row=r, column=start_c, value="")

        for c in range(1, 17):
            cell = ws.cell(row=r, column=c)
            cell.font = FONT_SMALL
            cell.fill = FILL_WHITE
            cell.alignment = ALIGN_LEFT
            cell.border = BORDER_GRID

    # --- 16行目: 注意事項（高さ 2.0cm = 56.7pt） ---
    ws.row_dimensions[16].height = 56.7
    ws.merge_cells("A16:B16")
    ws.merge_cells("C16:P16")
    ws["A16"] = "注意事項"
    ws["C16"] = ""
    for c in range(1, 17):
        cell = ws.cell(row=16, column=c)
        is_label = c in (1, 2)
        cell.font = FONT_SMALL_BOLD if is_label else FONT_SMALL
        cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
        cell.alignment = ALIGN_CENTER if is_label else ALIGN_TOP_LEFT
        cell.border = BORDER_GRID

    # --- 17行目: 作業日 枠（手書き記入用・年月日空欄） ---
    ws.row_dimensions[17].height = 17.5
    ws.merge_cells("A17:B17")
    ws.merge_cells("C17:P17")
    ws["A17"] = "作業日"
    ws["C17"] = "     年    月    日"
    for c in range(1, 17):
        cell = ws.cell(row=17, column=c)
        is_label = c in (1, 2)
        cell.font = FONT_SMALL_BOLD if is_label else FONT_SMALL
        cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
        cell.alignment = ALIGN_CENTER if is_label else ALIGN_LEFT
        cell.border = BORDER_GRID

    # --- 18行目: 表組② ヘッダー（現況のまま） ---
    ws.row_dimensions[18].height = 16.5
    sub_headers = ["No", "担当", "折メーター", "折損"] * 4
    for idx, text in enumerate(sub_headers):
        c = idx + 1
        cell = ws.cell(row=18, column=c, value=text)
        cell.font = FONT_SMALL_BOLD
        cell.fill = FILL_LABEL_BG
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_GRID

    # --- 19〜28行目: 表組②（手書き用空欄・現況のまま・高さ 1.2cm = 34.0pt） ---
    for r_offset in range(10):
        r = 19 + r_offset
        ws.row_dimensions[r].height = 34.0
        for block_idx in range(4):
            base_col = block_idx * 4 + 1
            no_val = r_offset + 1 + (block_idx * 10)

            c_no = ws.cell(row=r, column=base_col, value=no_val)
            c_no.font = FONT_SMALL_BOLD
            c_no.fill = FILL_LABEL_BG
            c_no.alignment = ALIGN_CENTER
            c_no.border = BORDER_GRID

            for sub_c in range(1, 4):
                cell = ws.cell(row=r, column=base_col + sub_c, value="")
                cell.font = FONT_SMALL
                cell.fill = FILL_WHITE
                cell.alignment = ALIGN_CENTER
                cell.border = BORDER_GRID

    # --- 29行目: 注意事項（高さ 2.0cm = 56.7pt） ---
    ws.row_dimensions[29].height = 56.7
    ws.merge_cells("A29:B29")
    ws.merge_cells("C29:P29")
    ws["A29"] = "注意事項"
    ws["C29"] = ""
    for c in range(1, 17):
        cell = ws.cell(row=29, column=c)
        is_label = c in (1, 2)
        cell.font = FONT_SMALL_BOLD if is_label else FONT_SMALL
        cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
        cell.alignment = ALIGN_CENTER if is_label else ALIGN_TOP_LEFT
        cell.border = BORDER_GRID

    # 1シート内改ページ（29行目の後で2頁目へ）
    ws.row_breaks.append(Break(id=29))

    # ==========================================
    # 2頁目: 貼り込み表
    # ==========================================
    ws.row_dimensions[30].height = 24
    ws.merge_cells("A30:P30")
    ws["A30"] = "貼 り 込 み 表"
    for c in range(1, 17):
        cell = ws.cell(row=30, column=c)
        cell.font = FONT_TITLE_MENKAKE
        cell.fill = FILL_HEADER_GRAY
        cell.alignment = ALIGN_CENTER
    _apply_box_border(ws, 1, 30, 16, 30)

    ws.row_dimensions[31].height = 5

    # --- 32行目: 書名・部数・作業日 ---
    ws.row_dimensions[32].height = 22
    ws.merge_cells("A32:B32")
    ws.merge_cells("C32:G32")
    ws.merge_cells("I32:J32")
    ws.merge_cells("K32:L32")
    ws.merge_cells("M32:P32")

    ws["A32"] = "書名"
    ws["C32"] = _safe_str(data.get("field-13"))
    ws["H32"] = "部数"
    ws["I32"] = f"{_safe_str(data.get('field-4', '500'))} 部"
    ws["K32"] = "作業日"
    ws["M32"] = "   年   月   日"

    for c in range(1, 17):
        cell = ws.cell(row=32, column=c)
        is_label = c in (1, 2, 8, 11, 12)
        cell.font = FONT_SMALL_BOLD if is_label else FONT_SMALL
        cell.fill = FILL_LABEL_BG if is_label else FILL_WHITE
        if is_label:
            cell.alignment = ALIGN_CENTER
        elif c in (9, 10):
            cell.alignment = ALIGN_RIGHT
        else:
            cell.alignment = ALIGN_LEFT
        cell.border = BORDER_GRID

    ws.row_dimensions[33].height = 5

    # --- 34行目: 貼り込み表 ヘッダー ---
    ws.row_dimensions[34].height = 22
    ws.merge_cells("A34:E34")
    ws.merge_cells("F34:G34")
    ws.merge_cells("H34:I34")
    ws.merge_cells("J34:K34")
    ws.merge_cells("L34:P34")

    ws["A34"] = "貼り込み箇所"
    ws["F34"] = "本数"
    ws["H34"] = "半端"
    ws["J34"] = "ヤレ枚数"
    ws["L34"] = "備考"

    for c in range(1, 17):
        cell = ws.cell(row=34, column=c)
        cell.font = FONT_SMALL_BOLD
        cell.fill = FILL_LABEL_BG
        cell.alignment = ALIGN_CENTER
        cell.border = BORDER_GRID

    # --- 35〜58行目: 貼り込み表 24行（折〜折 中央寄せ） ---
    for i in range(24):
        r = 35 + i
        ws.row_dimensions[r].height = 24
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
        ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=7)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        ws.merge_cells(start_row=r, start_column=10, end_row=r, end_column=11)
        ws.merge_cells(start_row=r, start_column=12, end_row=r, end_column=16)

        ws.cell(row=r, column=1, value="    折 〜     折")
        ws.cell(row=r, column=6, value="本")
        ws.cell(row=r, column=8, value="本")
        ws.cell(row=r, column=10, value="枚")
        ws.cell(row=r, column=12, value="")

        for c in range(1, 17):
            cell = ws.cell(row=r, column=c)
            cell.font = FONT_NORMAL
            cell.fill = FILL_WHITE
            if 1 <= c <= 5:
                cell.alignment = ALIGN_CENTER
            elif 6 <= c <= 11:
                cell.alignment = ALIGN_RIGHT
            else:
                cell.alignment = ALIGN_LEFT
            cell.border = BORDER_GRID

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ==========================================
# 2-B. 面掛表・貼り込み表 PDF生成（2ページ構成・A4縦収まり調整）
# ==========================================
def generate_menkake_pdf(data):
    pdf_path = "temp_menkake.pdf"
    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=A4,
        rightMargin=8 * mm, leftMargin=8 * mm,
        topMargin=3.5 * mm, bottomMargin=3.5 * mm
    )

    title_style = ParagraphStyle('Title14', fontName=DEFAULT_FONT, fontSize=13.5, leading=16, alignment=1)
    normal_style = ParagraphStyle('Normal7', fontName=DEFAULT_FONT, fontSize=7, leading=8.8)
    bold_style = ParagraphStyle('Bold7', fontName=DEFAULT_FONT, fontSize=7, leading=8.8)
    center_style = ParagraphStyle('Center7', fontName=DEFAULT_FONT, fontSize=7, leading=8.8, alignment=1)
    right_style = ParagraphStyle('Right7', fontName=DEFAULT_FONT, fontSize=7, leading=8.8, alignment=2)
    center_65pt = ParagraphStyle('Center65', fontName=DEFAULT_FONT, fontSize=6.5, leading=8, alignment=1)
    bold_center_65pt = ParagraphStyle('BoldCenter65', fontName=DEFAULT_FONT, fontSize=6.5, leading=8, alignment=1)

    story = []

    # --- 1頁目: 面掛表 ---
    story.append(Paragraph("<b>面掛表</b>", title_style))
    story.append(Spacer(1, 0.8 * mm))

    top_info_data = [
        [
            Paragraph(f"<b>受注番号:</b> {_safe_str(data.get('field-2'))}", normal_style),
            Paragraph(f"<b>出版社名:</b> {_safe_str(data.get('field-8'))}", normal_style),
        ]
    ]
    top_info_table = Table(top_info_data, colWidths=[97 * mm, 97 * mm], rowHeights=[4.8 * mm])
    top_info_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 1.5),
    ]))
    story.append(top_info_table)
    story.append(Spacer(1, 0.8 * mm))

    book_title_data = [[Paragraph(f"<b>書名:</b> {_safe_str(data.get('field-13'))}", bold_style)]]
    book_title_table = Table(book_title_data, colWidths=[194 * mm], rowHeights=[5 * mm])
    book_title_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 1.5),
    ]))
    story.append(book_title_table)
    story.append(Spacer(1, 0.8 * mm))

    spec_data = [
        [
            Paragraph(f"<b>総頁数:</b> {_safe_str(data.get('field-82'))}", normal_style),
            Paragraph(f"<b>折数:</b> {_format_fold_breakdown(data)}", normal_style),
            Paragraph(f"<b>合計台数:</b> {_safe_str(data.get('field-35'))}", normal_style),
        ]
    ]
    spec_table = Table(spec_data, colWidths=[34 * mm, 124 * mm, 36 * mm], rowHeights=[5 * mm])
    spec_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 1.5),
    ]))
    story.append(spec_table)
    story.append(Spacer(1, 0.8 * mm))

    t1_headers = [
        Paragraph("<b>面掛</b>", bold_style),
        Paragraph("<b>備考</b>", bold_style),
        Paragraph("<b>面掛</b>", bold_style),
        Paragraph("<b>備考</b>", bold_style),
    ]
    t1_rows = [t1_headers]
    for _ in range(10):
        t1_rows.append([
            Paragraph("", normal_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
            Paragraph("", normal_style),
        ])

    t1_table = Table(
        t1_rows,
        colWidths=[42 * mm, 55 * mm, 42 * mm, 55 * mm],
        rowHeights=[4.8 * mm] + [9.0 * mm] * 10
    )
    t1_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
        ('PADDING', (0, 0), (-1, -1), 1.5),
    ]))
    story.append(t1_table)
    story.append(Spacer(1, 0.8 * mm))

    note_data1 = [[Paragraph("<b>注意事項:</b>", normal_style)]]
    note_table1 = Table(note_data1, colWidths=[194 * mm], rowHeights=[20.0 * mm])
    note_table1.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('PADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(note_table1)
    story.append(Spacer(1, 1.0 * mm))

    space_ym = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    work_date_str = f"<b>作業日:</b>{space_ym}年{space_ym}月{space_ym}日"
    work_date_data = [[Paragraph(work_date_str, normal_style)]]
    work_date_table = Table(work_date_data, colWidths=[194 * mm], rowHeights=[5.0 * mm])
    work_date_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 1.5),
    ]))
    story.append(work_date_table)
    story.append(Spacer(1, 0.8 * mm))

    sub_headers = [
        Paragraph("<b>No</b>", bold_center_65pt), Paragraph("<b>担当</b>", center_65pt), Paragraph("<b>折メーター</b>", center_65pt), Paragraph("<b>折損</b>", bold_center_65pt),
        Paragraph("<b>No</b>", bold_center_65pt), Paragraph("<b>担当</b>", center_65pt), Paragraph("<b>折メーター</b>", center_65pt), Paragraph("<b>折損</b>", bold_center_65pt),
        Paragraph("<b>No</b>", bold_center_65pt), Paragraph("<b>担当</b>", center_65pt), Paragraph("<b>折メーター</b>", center_65pt), Paragraph("<b>折損</b>", bold_center_65pt),
        Paragraph("<b>No</b>", bold_center_65pt), Paragraph("<b>担当</b>", center_65pt), Paragraph("<b>折メーター</b>", center_65pt), Paragraph("<b>折損</b>", bold_center_65pt),
    ]
    sub_rows = [sub_headers]
    for r in range(10):
        sub_rows.append([
            Paragraph(str(r + 1), center_65pt), Paragraph("", normal_style), Paragraph("", normal_style), Paragraph("", center_65pt),
            Paragraph(str(r + 11), center_65pt), Paragraph("", normal_style), Paragraph("", normal_style), Paragraph("", center_65pt),
            Paragraph(str(r + 21), center_65pt), Paragraph("", normal_style), Paragraph("", normal_style), Paragraph("", center_65pt),
            Paragraph(str(r + 31), center_65pt), Paragraph("", normal_style), Paragraph("", normal_style), Paragraph("", center_65pt),
        ])

    sub_col_widths = [
        7.5 * mm, 13 * mm, 19 * mm, 9 * mm,
        7.5 * mm, 13 * mm, 19 * mm, 9 * mm,
        7.5 * mm, 13 * mm, 19 * mm, 9 * mm,
        7.5 * mm, 13 * mm, 19 * mm, 9 * mm,
    ]
    sub_table = Table(
        sub_rows,
        colWidths=sub_col_widths,
        rowHeights=[4.8 * mm] + [12.0 * mm] * 10
    )
    sub_table.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('PADDING', (0, 0), (-1, -1), 1),
    ]))
    story.append(sub_table)
    story.append(Spacer(1, 0.8 * mm))

    note_data2 = [[Paragraph("<b>注意事項:</b>", normal_style)]]
    note_table2 = Table(note_data2, colWidths=[194 * mm], rowHeights=[20.0 * mm])
    note_table2.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('PADDING', (0, 0), (-1, -1), 2),
    ]))
    story.append(note_table2)

    # --- 2頁目: 貼り込み表 ---
    story.append(PageBreak())

    story.append(Paragraph("<b>貼り込み表</b>", title_style))
    story.append(Spacer(1, 2 * mm))

    work_date_pasting = f"<b>作業日:</b>{space_ym}年{space_ym}月{space_ym}日"

    pasting_info_data = [
        [
            Paragraph(f"<b>書名:</b> {_safe_str(data.get('field-13'))}", normal_style),
            Paragraph(f"<b>部数:</b> {_safe_str(data.get('field-4', '500'))} 部", normal_style),
            Paragraph(work_date_pasting, normal_style),
        ]
    ]
    pasting_info_table = Table(pasting_info_data, colWidths=[84 * mm, 45 * mm, 65 * mm], rowHeights=[6 * mm])
    pasting_info_table.setStyle(TableStyle([
        ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, -1), colors.whitesmoke),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 2.5),
    ]))
    story.append(pasting_info_table)
    story.append(Spacer(1, 2 * mm))

    pasting_headers2 = [
        Paragraph("<b>貼り込み箇所</b>", bold_style),
        Paragraph("<b>本数</b>", bold_style),
        Paragraph("<b>半端</b>", bold_style),
        Paragraph("<b>ヤレ枚数</b>", bold_style),
        Paragraph("<b>備考</b>", bold_style),
    ]
    pasting_rows2 = [pasting_headers2]

    nbsp14 = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
    pasting_cell_text = f"{nbsp14}折〜{nbsp14}折"

    for _ in range(24):
        pasting_rows2.append([
            Paragraph(pasting_cell_text, center_style),
            Paragraph("本", right_style),
            Paragraph("本", right_style),
            Paragraph("枚", right_style),
            Paragraph("", normal_style),
        ])

    pasting_table2 = Table(
        pasting_rows2,
        colWidths=[60 * mm, 24.5 * mm, 24.5 * mm, 24.5 * mm, 60.5 * mm],
        rowHeights=[9.2 * mm] * 25
    )
    pasting_table2.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BACKGROUND', (0, 0), (-1, 0), colors.whitesmoke),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 1), (0, -1), 'CENTER'),
        ('ALIGN', (1, 1), (3, -1), 'RIGHT'),
        ('PADDING', (0, 0), (-1, -1), 2),
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
            st.write("袋払い（field-71）・折数表記・2頁チェックリスト対応")
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
            st.markdown("#### 📋 面掛表・貼り込み表（全2頁）")
            st.write("面掛表・表組①(0.9cm)・表組②(1.2cm)・注意事項(2cm) ＆ 貼り込み表")
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