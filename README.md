# 個資掃描程式 (Excel PII Scanner)

掃描 Excel 中兩個欄位的個資內容，統計涵蓋哪些個資類型、各有多少筆（含重複），
在終端機印出統計並輸出一份 Excel 報表。

掃描類型：

| 類型 | 說明 | 方式 |
|------|------|------|
| `ID` | 身分證字號 | regex |
| `phone` | 電話 / 手機 | regex |
| `mail` | Email | regex |
| `acc` | 帳號 / 卡號 | regex |
| `address` | 地址 | regex |
| `name` | 姓名 | [CKIP](https://github.com/ckiplab/ckip-transformers) NER (PERSON) |

## 環境

- Python 3.12
- 安裝依賴：

```bash
pip install -r requirements.txt
```

> 姓名掃描使用 CKIP NER 模型，**首次執行會自動下載模型權重，需要對外網路**。

## 使用方式

1. 開啟 `pii_scan.py`，修改檔頂的設定（不使用命令列參數）：

```python
INPUT_FILE = "data.xlsx"    # 來源 Excel
COL1 = "欄位A"               # 第一個要掃描的欄位名稱
COL2 = "欄位B"               # 第二個要掃描的欄位名稱
OUTPUT_FILE = "pii_report.xlsx"
```

2. 執行：

```bash
python pii_scan.py
```

## 輸出

**終端機**：每種類型在兩個欄位的命中筆數、合計，最後是每欄總和與全表總計。

**Excel 報表 (`pii_report.xlsx`)**，三個工作表：

- `統計總表`：類型 × 欄位，含合計列/欄。
- `欄位總和`：每個欄位各類型小計與總和。
- `明細`：每一筆命中（欄位｜列號｜類型｜命中內容）。

> 計數方式為「命中次數（含重複）」：同一個值出現多次都會分別計入。
