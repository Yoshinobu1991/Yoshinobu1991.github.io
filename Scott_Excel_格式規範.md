# 🚨 Scott 專用 Excel 庫存清單更新規範 (CRITICAL RULES) 🚨

**目標檔案**：`剩餘庫存清單-Scott專用格式.xlsx`

⚠️ **【絕對禁止事項】** ⚠️
絕對、絕對、絕對**不可以使用 `pandas`** (`pd.read_excel` / `df.to_excel`) 來修改或存檔此檔案！
`pandas` 會將檔案中所有的底色、框線、字體設定以及第二個活頁簿完全抹除！
**只能使用 `openpyxl` 進行精準修改！**

---

## 📌 格式規範與更新指南

### 1. 第一頁：`Funko庫存清單` (更新售出狀態)
當商品售出時，請依照以下格式將該商品標示為售出：
* **【不】需要加上任何「售出」字眼**。
* **整列反灰**：A 到 E 欄的底色必須改為灰色（Hex: `D9D9D9`）。
* **黑色刪除線**：A 到 E 欄的文字必須加上**黑色刪除線**（Strike=True, Color="000000"），並盡量保留原本的字體大小與字型（例如微軟正黑體 11）。

**Python (`openpyxl`) 實作範例**：
```python
gray_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
for cell in row[:5]:
    cell.fill = gray_fill
    cell.font = Font(
        name=cell.font.name if cell.font else '微軟正黑體',
        size=cell.font.size if cell.font else 11,
        strike=True,
        color="000000"
    )
```

### 2. 第二頁：`銷售紀錄` (新增售出訂單)
當新增一筆售出訂單時，請將資料附加至最下方，並嚴格遵守以下格式：
* **欄位結構**：
  * **A 欄 (日期)**：格式必須為 `YYYY-MM-DD`（例如 `2026-08-30`）。**底色必須統一為紫色**（Hex: `CCC0DA`）。
  * **B 欄 (售出項目)**：格式為 `#編號 角色名稱`。若是多個商品，請在同一個儲存格內使用 `\n` 換行。**底色必須為純白色**（Hex: `FFFFFF`）。
  * **C 欄 (總金額)**：填入整筆訂單的總額數字。**底色必須為純白色**（Hex: `FFFFFF`）。
* **框線與對齊**：
  * 所有新增的儲存格 (A、B、C 欄) 必須加上**黑色細實線 (thin) 的全框線**。
  * 所有新增的儲存格必須設定為**「水平與垂直置中」**，並且設定**「自動換行 (wrap_text=True)」**。
  * 字體建議使用「微軟正黑體」大小 12。

**Python (`openpyxl`) 實作範例**：
```python
purple_fill = PatternFill(start_color="CCC0DA", end_color="CCC0DA", fill_type="solid")
white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
center_wrap = Alignment(horizontal='center', vertical='center', wrap_text=True)

# A欄: 紫色
cell_A.fill = purple_fill
cell_A.border = thin_border
cell_A.alignment = center_wrap

# B, C欄: 白色
for cell in [cell_B, cell_C]:
    cell.fill = white_fill
    cell.border = thin_border
    cell.alignment = center_wrap
```
