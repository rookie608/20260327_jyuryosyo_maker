#!/usr/bin/env python3
"""
========================================================
書留・特定記録郵便物等差出票 PDF生成スクリプト
========================================================

【概要】
  inputフォルダに置いたCSVファイルを読み込み、
  郵便局提出用の「書留・特定記録郵便物等差出票」をPDFで自動生成します。

【フォルダ構成】
  create_slip.py   ← このスクリプト
  input/           ← 処理したいCSVファイルをここに入れる
      参加権01_xxx.csv
      参加権02_yyy.csv
      ...
  output/          ← 生成されたPDFが保存される（自動作成）
      参加権01_xxx_slip.pdf
      参加権02_yyy_slip.pdf
      ...

【CSVファイルの形式】
  - 文字コード: UTF-8 または Shift-JIS（自動判定）
  - 必須カラム: 「氏名」を含む列名（例: 氏名、お名前、name など）
  - 例:
      氏名,その他の列
      田中太郎,...
      鈴木花子,...

【摘要欄の記載内容】
  ファイル名に「参加権XX」が含まれる場合、それを自動抽出して使用します。
  例) ファイル名「参加権A_リベ大.csv」→ 摘要欄「参加権A：0001」
  ※ 通し番号はCSV内の行順（1始まり、4桁ゼロ埋め）

【出力ルール】
  - 1ページあたり30件
  - 1CSVあたり1PDFを出力
  - 1000件を超える場合は1000件ごとに分割して出力
      例) 1647件 → _slip_001.pdf（1〜1000件）
                    _slip_002.pdf（1001〜1647件）

【実行方法】
  # 仮想環境を有効化（初回セットアップ後は毎回必要）
  source .venv/bin/activate

  # スクリプトを実行
  python3 create_slip.py

【初回セットアップ】
  python3 -m venv .venv
  source .venv/bin/activate
  pip install -r requirements.txt

【差出人住所】
  以下が固定で印字されます（変更する場合は SENDER_LINE1 / SENDER_LINE2 を編集）
  〒541-0053 大阪府大阪市中央区本町２丁目３−４ 6階
  就労継続支援B型事業所 リベ大スキルアップ工房
========================================================
"""

import os
import re
import sys
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# ─── 日本語フォント登録（明朝体） ────────────────────────
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiMin-W3'))
pdfmetrics.registerFont(UnicodeCIDFont('HeiseiKakuGo-W5'))

FONT = 'HeiseiMin-W3'

# ─── 差出人住所（固定） ───────────────────────────────────
SENDER_LINE1 = '〒541-0053 大阪府大阪市中央区本町２丁目３−４ 6階'
SENDER_LINE2 = '就労継続支援B型事業所 リベ大スキルアップ工房'

# ─── 件数設定 ─────────────────────────────────────────────
ITEMS_PER_PAGE = 30    # 1ページあたりの件数
MAX_PER_PDF    = 1000  # 1PDFあたりの最大件数（超えると分割）

# ─── ページサイズ・余白 ───────────────────────────────────
PAGE_W, PAGE_H = A4   # A4縦: 595 × 841 pt
ML = MR = MT = MB = 10 * mm
USABLE_W = PAGE_W - ML - MR

# ─── 縦レイアウト（上から順に高さを定義） ────────────────
TITLE_H     = 9 * mm   # タイトル行
NOTE_H      = 6 * mm   # (ご依頼主のご住所・お名前) テキスト行
ADDR_H      = 24 * mm  # 差出人住所ボックス
GAP_H       = 2 * mm   # 住所ボックスとテーブルの間隔
HDR_H       = 8 * mm   # テーブルヘッダー行
# 残りの高さを30行で均等分割
DATA_AREA_H = PAGE_H - MT - MB - TITLE_H - NOTE_H - ADDR_H - GAP_H - HDR_H
ROW_H       = DATA_AREA_H / ITEMS_PER_PAGE

# ─── 横レイアウト（slip.xlsx の列比率を再現） ────────────
# 元ファイルの列幅比率: A=22.75, B=4.25, C=22.75, D=9.875, E≈7, F+G=21.75
RATIO_TOTAL = 22.75 + 4.25 + 22.75 + 9.875 + 7 + 21.75
W_NAME = USABLE_W * (22.75 / RATIO_TOTAL)  # お届け先のお名前
W_SAMA = USABLE_W * (4.25  / RATIO_TOTAL)  # 様
W_INQ  = USABLE_W * (22.75 / RATIO_TOTAL)  # お問い合わせ番号
W_DMG  = USABLE_W * (9.875 / RATIO_TOTAL)  # 申出損害要償額
W_FEE  = USABLE_W * (7     / RATIO_TOTAL)  # 料金等
W_NOTE = USABLE_W - W_NAME - W_SAMA - W_INQ - W_DMG - W_FEE  # 摘要

# ─── 罫線の太さ ───────────────────────────────────────────
THIN  = 0.5  # 細線（行の区切り）
THICK = 1.5  # 太線（外枠・列の区切り）


# ─── 描画ユーティリティ ───────────────────────────────────

def hline(c, x1, x2, y, lw=THIN):
    """水平線を描画する"""
    c.setLineWidth(lw)
    c.line(x1, y, x2, y)


def vline(c, x, y1, y2, lw=THIN):
    """垂直線を描画する"""
    c.setLineWidth(lw)
    c.line(x, y1, x, y2)


def centered(c, text, cx, y, font=FONT, fs=8):
    """テキストをX方向に中央揃えで描画する"""
    c.setFont(font, fs)
    c.drawCentredString(cx, y, text)


# ─── 1ページ描画 ─────────────────────────────────────────

def draw_page(c, records):
    """
    1ページ分を描画する。
    records: list of (name, note_text) — 最大 ITEMS_PER_PAGE 件
    note_text は build_note() で組み立て済みの摘要文字列。
    """

    # Y座標の基準点を上から順に計算（reportlabは左下原点・上方向が正）
    y_title_top = PAGE_H - MT
    y_note_top  = y_title_top - TITLE_H
    y_addr_top  = y_note_top  - NOTE_H
    y_table_top = y_addr_top  - ADDR_H - GAP_H
    y_data_top  = y_table_top - HDR_H

    # 各列のX座標
    x0     = ML
    x_sama = x0     + W_NAME
    x_inq  = x_sama + W_SAMA
    x_dmg  = x_inq  + W_INQ
    x_fee  = x_dmg  + W_DMG
    x_note = x_fee  + W_FEE
    x_end  = x_note + W_NOTE

    # ── タイトル ──────────────────────────────────────────
    title_fs = 13
    title_y  = y_title_top - TITLE_H / 2 - title_fs * 0.35
    centered(c, '書留・特定記録郵便物等差出票', PAGE_W / 2, title_y, fs=title_fs)

    # ── (ご依頼主のご住所・お名前) ────────────────────────
    note_fs = 8
    note_y  = y_note_top - NOTE_H / 2 - note_fs * 0.35
    c.setFont(FONT, note_fs)
    c.drawString(x0, note_y, '（ご依頼主のご住所・お名前）')

    # ── 差出人住所ボックス ────────────────────────────────
    c.setLineWidth(THICK)
    c.rect(x0, y_addr_top - ADDR_H, USABLE_W, ADDR_H)

    addr_fs  = 9
    addr_pad = 2 * mm
    line1_y  = y_addr_top - ADDR_H * 0.38 - addr_fs * 0.35
    line2_y  = y_addr_top - ADDR_H * 0.68 - addr_fs * 0.35
    c.setFont(FONT, addr_fs)
    c.drawString(x0 + addr_pad, line1_y, SENDER_LINE1)
    c.drawString(x0 + addr_pad, line2_y, SENDER_LINE2)

    # 様（右端）
    sama_fs = 12
    sama_y  = y_addr_top - ADDR_H / 2 - sama_fs * 0.35
    c.setFont(FONT, sama_fs)
    c.drawRightString(x_end - 2 * mm, sama_y, '様')

    # ── テーブルヘッダー ──────────────────────────────────
    hdr_top = y_table_top
    hdr_bot = hdr_top - HDR_H
    hdr_mid = (hdr_top + hdr_bot) / 2

    def hdr_text(text, cx, fs=8):
        centered(c, text, cx, hdr_mid - fs * 0.35, fs=fs)

    hline(c, x0, x_end, hdr_top, THIN)
    # 名前列・摘要列の下辺は太線（元ファイルの書式に合わせる）
    hline(c, x0,    x_sama + W_SAMA, hdr_bot, THICK)
    hline(c, x_inq, x_note,          hdr_bot, THIN)
    hline(c, x_note, x_end,          hdr_bot, THICK)

    vline(c, x0,             hdr_top, hdr_bot, THIN)
    vline(c, x_sama + W_SAMA, hdr_top, hdr_bot, THIN)
    vline(c, x_inq,           hdr_top, hdr_bot, THIN)
    vline(c, x_dmg,           hdr_top, hdr_bot, THIN)
    vline(c, x_fee,           hdr_top, hdr_bot, THIN)
    vline(c, x_note,          hdr_top, hdr_bot, THIN)
    vline(c, x_end,           hdr_top, hdr_bot, THIN)

    hdr_text('お届け先のお名前', x0 + (W_NAME + W_SAMA) / 2)
    hdr_text('お問い合わせ番号', x_inq + W_INQ / 2)
    hdr_text('申出損害要償額',   x_dmg + W_DMG / 2, fs=7)
    hdr_text('料金等',           x_fee + W_FEE / 2)
    hdr_text('摘　要',           x_note + W_NOTE / 2)

    # ── データ行（30行分の枠を描画し、データがある行に氏名・摘要を記入） ──
    data_fs  = 9    # 氏名フォントサイズ
    note_fs  = 7.5  # 摘要フォントサイズ
    sama2_fs = 9    # 「様」フォントサイズ

    for i in range(ITEMS_PER_PAGE):
        row_top = y_data_top - i * ROW_H
        row_bot = row_top - ROW_H
        row_mid = (row_top + row_bot) / 2 - data_fs * 0.35
        is_last = (i == ITEMS_PER_PAGE - 1)

        # 水平線（最終行の下辺は太線）
        hline(c, x0, x_end, row_top, THIN)
        if is_last:
            hline(c, x0,    x_sama + W_SAMA, row_bot, THICK)
            hline(c, x_inq, x_note,          row_bot, THIN)
            hline(c, x_note, x_end,          row_bot, THICK)
        else:
            hline(c, x0, x_end, row_bot, THIN)

        # 縦線（名前列左右・摘要列左右は太線、それ以外は細線）
        vline(c, x0,              row_top, row_bot, THICK)  # 名前列 左
        vline(c, x_sama,          row_top, row_bot, THIN)   # 名前/様 区切り
        vline(c, x_sama + W_SAMA, row_top, row_bot, THICK)  # 名前列 右
        vline(c, x_inq + W_INQ,   row_top, row_bot, THIN)
        vline(c, x_dmg + W_DMG,   row_top, row_bot, THIN)
        vline(c, x_fee + W_FEE,   row_top, row_bot, THIN)
        vline(c, x_note,          row_top, row_bot, THICK)  # 摘要列 左
        vline(c, x_end,           row_top, row_bot, THICK)  # 摘要列 右

        # 「様」は全行に印字
        c.setFont(FONT, sama2_fs)
        c.drawCentredString(x_sama + W_SAMA / 2, row_mid, '様')

        # データがある行のみ氏名・摘要を記入
        if i < len(records):
            name, note_text = records[i]

            if name:
                c.setFont(FONT, data_fs)
                c.drawRightString(x_sama - 1 * mm, row_mid, name)

            if note_text:
                c.setFont(FONT, note_fs)
                c.drawString(
                    x_note + 1.5 * mm,
                    row_mid + data_fs * 0.35 - note_fs * 0.35,
                    note_text,
                )


# ─── ファイル名から参加権ラベルを抽出 ────────────────────

def extract_sankaken(filename):
    """
    ファイル名から「参加権XX」を抽出して返す。
    例) '参加権A_リベ大.csv' → '参加権A'
    見つからない場合は None を返す（摘要欄の先頭ラベルを省略するため）。
    """
    base = os.path.splitext(os.path.basename(filename))[0]
    m = re.search(r'参加権[^\s_　/\\]+', base)
    return m.group(0) if m else None


def clean_name(name):
    """
    名前末尾の「様」と前後スペースを取り除く。
    CSV の名前列に「山田 太郎 様」と入っている場合に使用。
    テンプレート側で「様」を自動印字するため重複を防ぐ。
    """
    return re.sub(r'\s*様\s*$', '', str(name)).strip()


def build_note(sankaken_label, tsushi_val, uketsuke_val):
    """
    摘要欄の文字列を組み立てる。

    組み立てルール:
      - ファイル名に「参加権XX」がある場合 → 先頭に追加
      - 通し番号（CSV列 or 行連番）→ 必ず含める
      - 受付番号（CSV列がある場合のみ）→ 末尾に追加

    例）
      参加権A CSV（受付番号なし）: '参加権A：0001'
      このCSV（参加権なし、受付番号あり）: '8：000-009'
      両方ある場合: '参加権A：15：122267-2108'
    """
    parts = []
    if sankaken_label:
        parts.append(sankaken_label)
    parts.append(tsushi_val)
    if uketsuke_val:
        parts.append(uketsuke_val)
    return '：'.join(parts)


# ─── PDF生成 ─────────────────────────────────────────────

def generate_pdf(records, output_path):
    """
    records をページ分割してPDFに書き出す。
    records    : list of (name, note_text)
    output_path: 出力先PDFパス
    """
    c = canvas.Canvas(output_path, pagesize=A4)
    pages = [records[i:i + ITEMS_PER_PAGE] for i in range(0, len(records), ITEMS_PER_PAGE)]
    for page_records in pages:
        draw_page(c, page_records)
        c.showPage()
    c.save()
    print(f'  -> {output_path}  ({len(records)}件, {len(pages)}ページ)')


# ─── CSV処理 ─────────────────────────────────────────────

def process_csv(csv_path):
    """
    CSVを1ファイル読み込み、PDFを生成する。
    1000件を超える場合は1000件ごとに分割して複数のPDFを出力する。
    """
    print(f'読み込み: {csv_path}')

    # 文字コードを自動判定（UTF-8 → Shift-JIS の順で試行）
    try:
        df = pd.read_csv(csv_path, encoding='utf-8-sig')
    except UnicodeDecodeError:
        df = pd.read_csv(csv_path, encoding='shift_jis')

    # ── カラム検索 ────────────────────────────────────────
    # 名前: '名前' → '氏名' → 'name' の優先順で検索
    name_col = next((c for c in df.columns if '名前' in c or '氏名' in c or 'name' in c.lower()), None)
    if name_col is None:
        print(f'  [エラー] 名前カラムが見つかりません。カラム一覧: {list(df.columns)}')
        return

    # 通し番号列（あれば使う、なければ行の連番で代用）
    tsushi_col = next((c for c in df.columns if c == '通し番号' or ('通し' in c and '番号' in c)), None)

    # 受付番号列（あれば摘要末尾に追加）
    uketsuke_col = next((c for c in df.columns if c == '受付番号' or ('受付' in c and '番号' in c)), None)

    print(f'  使用カラム: 名前={name_col!r}, 通し番号={tsushi_col!r}, 受付番号={uketsuke_col!r}')

    # ── レコード組み立て ──────────────────────────────────
    sankaken_label = extract_sankaken(csv_path)
    print(f'  参加権ラベル: {sankaken_label!r}')

    records = []
    for i, (_, row) in enumerate(df.iterrows()):
        # 氏名（末尾の「様」を除去）
        name = clean_name(row[name_col]) if pd.notna(row[name_col]) else ''

        # 通し番号: CSV列があればそれを使用、なければ行連番（4桁ゼロ埋め）
        if tsushi_col and pd.notna(row[tsushi_col]):
            tsushi_val = str(row[tsushi_col]).strip()
        else:
            tsushi_val = str(i + 1).zfill(4)

        # 受付番号: CSV列があれば取得
        if uketsuke_col and pd.notna(row[uketsuke_col]):
            uketsuke_val = str(row[uketsuke_col]).strip()
        else:
            uketsuke_val = None

        note = build_note(sankaken_label, tsushi_val, uketsuke_val)
        records.append((name, note))

    base    = os.path.splitext(os.path.basename(csv_path))[0]
    out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')

    # 1000件超は分割出力
    chunks = [records[i:i + MAX_PER_PDF] for i in range(0, len(records), MAX_PER_PDF)]
    for idx, chunk in enumerate(chunks):
        if len(chunks) == 1:
            out_path = os.path.join(out_dir, f'{base}_slip.pdf')
        else:
            out_path = os.path.join(out_dir, f'{base}_slip_{idx + 1:03d}.pdf')
        generate_pdf(chunk, out_path)


# ─── エントリーポイント ───────────────────────────────────

def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_dir  = os.path.join(script_dir, 'input')
    output_dir = os.path.join(script_dir, 'output')

    if not os.path.isdir(input_dir):
        print(f'[エラー] inputフォルダが見つかりません: {input_dir}')
        print('スクリプトと同じ階層に input/ フォルダを作成し、CSVを入れてください。')
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)

    csv_files = sorted(f for f in os.listdir(input_dir) if f.lower().endswith('.csv'))
    if not csv_files:
        print(f'[エラー] inputフォルダにCSVファイルがありません: {input_dir}')
        sys.exit(1)

    print(f'inputフォルダ : {input_dir}')
    print(f'outputフォルダ: {output_dir}')
    print(f'対象CSV       : {len(csv_files)}件\n')

    for fname in csv_files:
        process_csv(os.path.join(input_dir, fname))


if __name__ == '__main__':
    main()
