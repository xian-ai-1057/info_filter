# -*- coding: utf-8 -*-
"""個資掃描程式 (Excel PII Scanner)

讀取 Excel 中兩個欄位，掃描其中的個資（身分證、電話、Email、帳號/卡號、
地址，以及使用 CKIP NER 的姓名），統計各類型命中筆數（含重複），
並在終端機印出統計、同時輸出一份 Excel 報表。

執行方式：python pii_scan.py
"""

import re

import pandas as pd

# ───────────── 設定（要改就改這幾行） ─────────────
INPUT_FILE = "data.xlsx"        # 來源 Excel 路徑
COL1 = "欄位A"                   # 第一個要掃描的欄位名稱
COL2 = "欄位B"                   # 第二個要掃描的欄位名稱
OUTPUT_FILE = "pii_report.xlsx"  # 報表輸出路徑
# ────────────────────────────────────────────────

# regex 個資樣式（原樣沿用）
patterns = {
    'ID':      re.compile(r'(?:[a-zA-Z][0-9]{9})|(?:[a-zA-Z]{2}[0-9]{8})'),
    'phone':   re.compile(r'([(]?\+?(0[2-9]|0[2-9]\d{2}|886[2-9]|037|049|089|0800)[)]?-?(\d{3,4})-?(\d{3,4}))'),
    'mail':    re.compile(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'),
    'acc':     re.compile(r'(?:\d{4}?[-\s]?){2}\d{2,4}?[-\s]?\d{4}|(?:\d{4}?[-\s]?){3}\d{3,4}'),
    'address': re.compile(
                # (1) 縣市名稱可有可無
                r'(?:'
                    r'(?:台灣|臺灣|台北市|臺北市|新北市|桃園市|台中市|臺中市|台南市|臺南市|高雄市|基隆市|'
                    r'新竹市|嘉義市|苗栗縣|彰化縣|南投縣|雲林縣|屏東縣|宜蘭縣|花蓮縣|台東縣|臺東縣|澎湖縣|'
                    r'金門縣|連江縣'
                    r'台北|臺北|新北|桃園|台中|臺中|台南|臺南|高雄|基隆|'
                    r'新竹|嘉義|苗栗|彰化|南投|雲林|屏東|宜蘭|花蓮|台東|臺東|澎湖|'
                    r'金門|連江)'
                    # 若有縣市，後面可能跟「XX區 / XX鎮 / 里 / 村」等各種文字
                    r'[一-龥\d\-\s]*'
                r')?'

                # (2) 必須出現「街 / 大道 / 路」
                r'(?:[一-龥\d\-\s]{2}(?:街|大道|路)).*?'

                # (3) 至少一次出現「巷 / 弄 / 號 / 樓 / 室」
                r'(?:[一-龥\d\-\s~之]*'
                r'(?:(?:巷|弄|樓|室)|(?<!手機)(?<!電話)(?<!網路)(?<!門)號)'
                r')+'
            )
}

# 統計時固定的類型順序（regex 類型 + 姓名）
TYPES = list(patterns.keys()) + ['name']


def scan_regex(cells):
    """掃描一欄的所有 cell，回傳 {類型: 命中次數} 與明細清單。

    cells: list of (excel_row, text)
    明細: list of dict(欄位 由呼叫端補上)
    """
    counts = {t: 0 for t in patterns}
    details = []
    for row, text in cells:
        for ptype, pat in patterns.items():
            for m in pat.finditer(text):
                counts[ptype] += 1
                details.append({'列號': row, '類型': ptype, '命中內容': m.group(0)})
    return counts, details


def scan_names(cells, ner):
    """用 CKIP NER 批次掃描姓名(PERSON)，回傳命中次數與明細。"""
    texts = [text for _, text in cells]
    rows = [row for row, _ in cells]
    count = 0
    details = []
    results = ner(texts)
    for row, entities in zip(rows, results):
        for ent in entities:
            if ent.ner == 'PERSON':
                count += 1
                details.append({'列號': row, '類型': 'name', '命中內容': ent.word})
    return count, details


def get_cells(df, col):
    """取出欄位中非空的 (excel_row, text)。excel_row 對應原檔列號(含標題列)。"""
    cells = []
    for i, val in enumerate(df[col]):
        if pd.isna(val):
            continue
        text = str(val).strip()
        if text:
            cells.append((i + 2, text))  # +2: 標題列 + 0-based
    return cells


def main():
    df = pd.read_excel(INPUT_FILE)

    for col in (COL1, COL2):
        if col not in df.columns:
            raise SystemExit(
                f"找不到欄位「{col}」，實際欄位為：{list(df.columns)}\n"
                f"請修改 pii_scan.py 頂端的 COL1 / COL2 設定。"
            )

    # 載入 CKIP NER 模型（首次執行會自動下載權重，需網路）
    print("載入 CKIP NER 模型中…（首次執行需下載權重）")
    from ckip_transformers.nlp import CkipNerChunker
    ner = CkipNerChunker(model="bert-base")

    col_counts = {}      # {欄位名: {類型: 次數}}
    all_details = []     # 完整明細

    for col in (COL1, COL2):
        cells = get_cells(df, col)
        counts, details = scan_regex(cells)
        name_count, name_details = scan_names(cells, ner)
        counts['name'] = name_count

        for d in details + name_details:
            d['欄位'] = col
            all_details.append(d)

        col_counts[col] = counts

    print_report(col_counts)
    write_report(col_counts, all_details)
    print(f"\n報表已輸出：{OUTPUT_FILE}")


def print_report(col_counts):
    """終端機輸出：類型 × 欄位 + 合計，最後欄位總和與全表總計。"""
    print("\n================ 個資掃描統計 ================")
    header = f"{'類型':<10}{COL1:>10}{COL2:>10}{'合計':>10}"
    print(header)
    print("-" * len(header.encode('utf-8')))

    grand = 0
    for t in TYPES:
        a = col_counts[COL1][t]
        b = col_counts[COL2][t]
        total = a + b
        grand += total
        print(f"{t:<10}{a:>10}{b:>10}{total:>10}")

    sum_a = sum(col_counts[COL1].values())
    sum_b = sum(col_counts[COL2].values())
    print("-" * len(header.encode('utf-8')))
    print(f"{'欄位總和':<10}{sum_a:>10}{sum_b:>10}{grand:>10}")
    print(f"\n全表個資命中總筆數：{grand}")
    print("=============================================")


def write_report(col_counts, all_details):
    """輸出 Excel 報表：統計總表 / 欄位總和 / 明細。"""
    # 統計總表（類型 × 欄位 + 合計列/欄）
    rows = []
    for t in TYPES:
        a = col_counts[COL1][t]
        b = col_counts[COL2][t]
        rows.append({'類型': t, COL1: a, COL2: b, '合計': a + b})
    summary = pd.DataFrame(rows)
    total_row = {
        '類型': '欄位總和',
        COL1: summary[COL1].sum(),
        COL2: summary[COL2].sum(),
        '合計': summary['合計'].sum(),
    }
    summary = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

    # 欄位總和（每欄各類型小計 + 總和）
    col_sum_rows = []
    for col in (COL1, COL2):
        sub = {'欄位': col}
        sub.update(col_counts[col])
        sub['總和'] = sum(col_counts[col].values())
        col_sum_rows.append(sub)
    col_sum = pd.DataFrame(col_sum_rows)

    # 明細
    if all_details:
        detail = pd.DataFrame(all_details)[['欄位', '列號', '類型', '命中內容']]
    else:
        detail = pd.DataFrame(columns=['欄位', '列號', '類型', '命中內容'])

    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        summary.to_excel(writer, sheet_name='統計總表', index=False)
        col_sum.to_excel(writer, sheet_name='欄位總和', index=False)
        detail.to_excel(writer, sheet_name='明細', index=False)


if __name__ == '__main__':
    main()
