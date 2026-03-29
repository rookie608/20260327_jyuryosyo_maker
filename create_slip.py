#!/usr/bin/env python3
"""
書留・特定記録郵便物等差出票 生成スクリプト

Usage:
    python create_slip.py <csv_file> [<csv_file> ...]

CSV必須カラム: 氏名, 受付番号
出力: <csv_basename>_slip.pdf (1000件超は _slip_001.pdf, _slip_002.pdf ...)
"""

import os
import re
import sys
import math
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))

FONT = 'HeiseiMin-W3'

SENDER_LINE1 = '〒541-0053 大阪府大阪市中央区本町２丁目３−４ 6階'
SENDER_LINE2 = '就労継続支援B型事業所 リベ大スキルアップ工房'

ITEMS_PER_PAGE = 30
MAX_PER_PDF = 1000

PAGE_W, PAGE_H = A4
ML = 10 * mm
MR = 10 * mm
MT = 10 * mm
MB = 10 * mm
USABLE_W = PAGE_W - ML - MR

# ─── 縦レイアウト(mm) ───────────────────────────
TITLE_H     = 9 * mm
NOTE_H      = 6 * mm
ADDR_H      = 24 * mm
GAP_H       = 2 * mm
HDR_H       = 8 * mm
# 残りを30行で均等分割
DATA_AREA_H = PAGE_H - MT - MB - TITLE_H - NOTE_H - ADDR_H - GAP_H - HDR_H
ROW_H       = DATA_AREA_H / ITEMS_PER_PAGE

# ─── 横レイアウト ────────────────────────────────
# slip.xlsx の列比率: A=22.75, B=4.25, C=22.75, D=9.875, E≈7, F+G=10.875*2
RATIO_TOTAL = 22.75 + 4.25 + 22.75 + 9.875 + 7 + 21.75
W_NAME  = USABLE_W * (22.75 / RATIO_TOTAL)
W_SAMA  = USABLE_W * (4.25  / RATIO_TOTAL)
W_INQ   = USABLE_W * (22.75 / RATIO_TOTAL)
W_DMG   = USABLE_W * (9.875 / RATIO_TOTAL)
W_FEE   = USABLE_W * (7     / RATIO_TOTAL)
W_NOTE  = USABLE_W - W_NAME - W_SAMA - W_INQ - W_DMG - W_FEE

THIN  = 0.5
THICK = 1.5


def hline(c, x1, x2, y, lw=THIN):
    c.setLineWidth(lw)
    c.line(x1, y, x2, y)


def vline(c, x, y1, y2, lw=THIN):
    c.setLineWidth(lw)
    c.line(x, y1, x, y2)


def centered(c, text, cx, y, font=FONT, fs=8):
    c.setFont(font, fs)
    c.drawCentredString(cx, y, text)


def draw_page(c, records, sankaken_label):
    """1ページ分を描画する。records: list of (name, tsushi_no)"""

    # ─── Y座標の基準点を上から順に計算 ───────────────
    y_title_top   = PAGE_H - MT
    y_note_top    = y_title_top  - TITLE_H
    y_addr_top    = y_note_top   - NOTE_H
    y_table_top   = y_addr_top   - ADDR_H - GAP_H
    y_data_top    = y_table_top  - HDR_H

    x0 = ML
    x_sama   = x0 + W_NAME
    x_inq    = x_sama + W_SAMA
    x_dmg    = x_inq  + W_INQ
    x_fee    = x_dmg  + W_DMG
    x_note   = x_fee  + W_FEE
    x_end    = x_note + W_NOTE

    # ─── タイトル ─────────────────────────────────────
    title_fs = 13
    title_y  = y_title_top - TITLE_H / 2 - title_fs * 0.35
    centered(c, '書留・特定記録郵便物等差出票', PAGE_W / 2, title_y, fs=title_fs)

    # ─── (ご依頼主のご住所・お名前) ───────────────────
    note_fs = 8
    note_y  = y_note_top - NOTE_H / 2 - note_fs * 0.35
    c.setFont(FONT, note_fs)
    c.drawString(x0, note_y, '（ご依頼主のご住所・お名前）')

    # ─── 住所ボックス ─────────────────────────────────
    # 外枠 thick
    c.setLineWidth(THICK)
    c.rect(x0, y_addr_top - ADDR_H, USABLE_W, ADDR_H)

    # 住所テキスト (左に2mm余白)
    addr_fs = 9
    addr_pad = 2 * mm
    line1_y = y_addr_top - ADDR_H * 0.38 - addr_fs * 0.35
    line2_y = y_addr_top - ADDR_H * 0.68 - addr_fs * 0.35
    c.setFont(FONT, addr_fs)
    c.drawString(x0 + addr_pad, line1_y, SENDER_LINE1)
    c.drawString(x0 + addr_pad, line2_y, SENDER_LINE2)

    # 様 (右端)
    sama_fs = 12
    sama_y  = y_addr_top - ADDR_H / 2 - sama_fs * 0.35
    c.setFont(FONT, sama_fs)
    c.drawRightString(x_end - 2 * mm, sama_y, '様')

    # ─── テーブルヘッダー ─────────────────────────────
    hdr_top = y_table_top
    hdr_bot = hdr_top - HDR_H
    hdr_mid = (hdr_top + hdr_bot) / 2

    def hdr_text(text, cx, fs=8):
        centered(c, text, cx, hdr_mid - fs * 0.35, fs=fs)

    # 外枠
    c.setLineWidth(THIN)
    # 上辺
    hline(c, x0, x_end, hdr_top, THIN)
    # 下辺: 名前列とノート列はTHICK
    hline(c, x0, x_sama + W_SAMA, hdr_bot, THICK)
    hline(c, x_inq, x_note, hdr_bot, THIN)
    hline(c, x_note, x_end, hdr_bot, THICK)
    # 縦線
    vline(c, x0,    hdr_top, hdr_bot, THIN)
    vline(c, x_sama + W_SAMA, hdr_top, hdr_bot, THIN)  # 名前列右
    vline(c, x_inq, hdr_top, hdr_bot, THIN)
    vline(c, x_dmg, hdr_top, hdr_bot, THIN)
    vline(c, x_fee, hdr_top, hdr_bot, THIN)
    vline(c, x_note, hdr_top, hdr_bot, THIN)
    vline(c, x_end, hdr_top, hdr_bot, THIN)

    hdr_text('お届け先のお名前', x0 + (W_NAME + W_SAMA) / 2)
    hdr_text('お問い合わせ番号', x_inq + W_INQ / 2)
    hdr_text('申出損害要償額',   x_dmg + W_DMG / 2, fs=7)
    hdr_text('料金等',           x_fee + W_FEE / 2)
    hdr_text('摘　要',           x_note + W_NOTE / 2)

    # ─── データ行 ─────────────────────────────────────
    data_fs  = 9
    note_fs  = 7.5
    sama2_fs = 9

    for i in range(ITEMS_PER_PAGE):
        row_top = y_data_top - i * ROW_H
        row_bot = row_top - ROW_H
        row_mid = (row_top + row_bot) / 2 - data_fs * 0.35
        is_last = (i == ITEMS_PER_PAGE - 1)

        # 水平線
        hline(c, x0,    x_end, row_top, THIN)
        if is_last:
            hline(c, x0,    x_sama + W_SAMA, row_bot, THICK)
            hline(c, x_inq, x_note,           row_bot, THIN)
            hline(c, x_note, x_end,            row_bot, THICK)
        else:
            hline(c, x0, x_end, row_bot, THIN)

        # 縦線
        vline(c, x0,    row_top, row_bot, THICK)   # 名前左
        vline(c, x_sama, row_top, row_bot, THIN)   # 名前/様の区切り
        vline(c, x_sama + W_SAMA, row_top, row_bot, THICK)  # 名前右
        vline(c, x_inq + W_INQ,   row_top, row_bot, THIN)
        vline(c, x_dmg + W_DMG,   row_top, row_bot, THIN)
        vline(c, x_fee + W_FEE,   row_top, row_bot, THIN)
        vline(c, x_note, row_top, row_bot, THICK)  # ノート左
        vline(c, x_end,  row_top, row_bot, THICK)  # ノート右

        # 様
        c.setFont(FONT, sama2_fs)
        c.drawCentredString(x_sama + W_SAMA / 2, row_mid, '様')

        # 氏名
        if i < len(records):
            name, tsushi_no = records[i]
            if name:
                c.setFont(FONT, data_fs)
                # 右寄せ（様の左に余白）
                c.drawRightString(x_sama - 1 * mm, row_mid, name)

            # 摘要
            if tsushi_no:
                note_text = f'{sankaken_label}：{tsushi_no}'
                c.setFont(FONT, note_fs)
                c.drawString(x_note + 1.5 * mm,
                             row_mid + data_fs * 0.35 - note_fs * 0.35,
                             note_text)


def extract_sankaken(filename):
    """ファイル名から「参加権XX」を抽出。見つからなければベース名を返す。"""
    base = os.path.splitext(os.path.basename(filename))[0]
    m = re.search(r'参加権[^\s_　/\\]+', base)
    return m.group(0) if m else base


def generate_pdf(records, output_path, sankaken_label):
    c = canvas.Canvas(output_path, pagesize=A4)
    pages = [records[i:i + ITEMS_PER_PAGE] for i in range(0, len(records), ITEMS_PER_PAGE)]
    for page_records in pages:
        draw_page(c, page_records, sankaken_label)
        c.showPage()
    c.save()
    print(f'  -> {output_path}  ({len(records)}件, {len(pages)}ページ)')


def process_csv(csv_path):
    print(f'読み込み: {csv_path}')
    try:
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
    except UnicodeDecodeError:
        df = pd.read_csv(csv_path, encoding='shift_jis')

    # カラム名の候補を検索
    name_col = next((c for c in df.columns if '氏名' in c or 'name' in c.lower()), None)

    if name_col is None:
        print(f'  [エラー] 氏名カラムが見つかりません。カラム一覧: {list(df.columns)}')
        return

    print(f'  使用カラム: 氏名={name_col!r}')

    records = [
        (
            str(row[name_col]).strip() if pd.notna(row[name_col]) else '',
            str(i + 1).zfill(4),  # 通し番号（1始まり、4桁ゼロ埋め）
        )
        for i, (_, row) in enumerate(df.iterrows())
    ]

    sankaken_label = extract_sankaken(csv_path)
    print(f'  参加権ラベル: {sankaken_label!r}')

    base    = os.path.splitext(os.path.basename(csv_path))[0]
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')

    chunks = [records[i:i + MAX_PER_PDF] for i in range(0, len(records), MAX_PER_PDF)]
    for idx, chunk in enumerate(chunks):
        if len(chunks) == 1:
            out_path = os.path.join(out_dir, f'{base}_slip.pdf')
        else:
            out_path = os.path.join(out_dir, f'{base}_slip_{idx + 1:03d}.pdf')
        generate_pdf(chunk, out_path, sankaken_label)


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir  = os.path.join(script_dir, 'input')
    output_dir = os.path.join(script_dir, 'output')

    if not os.path.isdir(input_dir):
        print(f'inputフォルダが見つかりません: {input_dir}')
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)

    csv_files = sorted(
        f for f in os.listdir(input_dir)
        if f.lower().endswith('.csv')
    )
    if not csv_files:
        print(f'inputフォルダにCSVファイルがありません: {input_dir}')
        sys.exit(1)

    print(f'inputフォルダ: {input_dir}')
    print(f'outputフォルダ: {output_dir}')
    print(f'対象CSV: {len(csv_files)}件\n')
    for fname in csv_files:
        process_csv(os.path.join(input_dir, fname))


if __name__ == '__main__':
    main()
