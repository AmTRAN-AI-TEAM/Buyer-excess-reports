# buyer-excess-reports

從 `BOM.xlsx` 與含有 `shortage` 分頁的 Excel，自動產出可用版 `Excess` 報表。

目前第一版目標是穩定產出 `Excess` sheet。價格、供應商、分類、呆料註記、前版金額等缺少明確來源的欄位會留空或為 0。

## 資料夾

```text
buyer-excess-reports/
  input/
  output/
  src/
```

## 安裝

```bash
cd ~/buyer-excess-reports
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## 使用

把輸入檔放進 `input/`：

```text
input/BOM.xlsx
input/你的shortage檔.xlsx
```

shortage 檔內只要有分頁名稱包含 `shortage`，程式就會使用該分頁。輸入檔不需要有 `Excess` sheet。

執行：

```bash
python3 src/main.py
```

輸出：

```text
output/excess_report.xlsx
```

## 目前產出邏輯

- 讀取 `shortage` 的 Part No、Description、Planner、Buyer、LT、MOQ。
- 讀取 `shortage` 的 `Overshortage1`、`PO_REMAIN` 與第一欄 WO 外 demand。
- 若找不到 `Overshortage1`，會用 `OVER_SHORTAGE + HLD` 當替代。
- 從 BOM 的 G 欄判斷替代料：非 `*R*` 是主料，後續同一 BOM 區段連續 `*R*` 是替代料。
- 替代料不做跨機種/跨 BOM 區段的全域串聯；同一料號若在多個 BOM 區段出現，程式會挑選最適合目前 shortage 料號的一個區段。
- 已經被某一列 Excess 使用過的 shortage 料號，不會在後續列重複加總。
- 同一替代料群組最多輸出 `PartNo1` 到 `替代料7`。
- 數量欄直接寫入計算後的值，不依賴輸出檔內公式。

## 自訂路徑

也可以指定輸入與輸出：

```bash
python3 src/main.py --input-dir /path/to/input --output-file /path/to/excess_report.xlsx
```
