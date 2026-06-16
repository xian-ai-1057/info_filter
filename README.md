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
MODEL_PATH = ""             # 留空＝線上下載；填本機模型資料夾＝離線使用（見下節）
```

2. 執行：

```bash
python pii_scan.py
```

## 使用本機（離線）模型

若機器無法連上 `huggingface.co`，可先手動下載姓名 NER 模型，再用 `MODEL_PATH` 指向該資料夾。

**1. 在有網路的機器下載模型** `ckiplab/bert-base-chinese-ner`，任一方式：

```bash
# 方式 A：huggingface-cli
pip install -U "huggingface_hub[cli]"
hf download ckiplab/bert-base-chinese-ner --local-dir bert-base-chinese-ner

# 方式 B：git（需安裝 git-lfs）
git lfs install
git clone https://huggingface.co/ckiplab/bert-base-chinese-ner
```

下載後資料夾應包含模型與 tokenizer 檔，例如：
`config.json`、`pytorch_model.bin`（或 `model.safetensors`）、`vocab.txt`、`tokenizer_config.json`、`special_tokens_map.json`。

**2. 設定路徑**：把整個資料夾複製到目標機器，並在 `pii_scan.py` 設定：

```python
MODEL_PATH = r"C:\models\bert-base-chinese-ner"   # 指向你的模型資料夾
```

設定後程式會自動開啟離線模式（`HF_HUB_OFFLINE` / `TRANSFORMERS_OFFLINE`），執行時不再嘗試連網。

## 輸出

**終端機**：每種類型在兩個欄位的命中筆數、合計，最後是每欄總和與全表總計。

**Excel 報表 (`pii_report.xlsx`)**，三個工作表：

- `統計總表`：類型 × 欄位，含合計列/欄。
- `欄位總和`：每個欄位各類型小計與總和。
- `明細`：每一筆命中（欄位｜列號｜類型｜命中內容）。

> 計數方式為「命中次數（含重複）」：同一個值出現多次都會分別計入。
