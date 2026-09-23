# buyer-excess-reports

從 `BOM.xlsx` 與含有 `shortage` 分頁的 Excel，自動產出可用版 `Excess` 報表。

目前第一版目標是穩定產出 `Excess` sheet。價格、供應商、分類、呆料註記、前版金額等缺少明確來源的欄位會留空。

## 資料夾

```text
buyer-excess-reports/
  input/
    AVTC/
    RAKEN/
  output/
    AVTC/
    RAKEN/
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

把輸入檔放進對應客戶資料夾：

```text
input/RAKEN/BOM.xlsx
input/RAKEN/你的shortage檔.xlsx

input/AVTC/BOM.xlsx
input/AVTC/你的shortage檔.xlsx
```

shortage 檔內只要有分頁名稱包含 `shortage`，程式就會使用該分頁。輸入檔不需要有 `Excess` sheet。

執行：

```bash
python3 src/main.py
```

程式會先讓你選擇 `AVTC` 或 `RAKEN`，再讀取對應的 `input/<客戶>/`，並在執行時顯示 `tqdm` 進度條。

輸出：

```text
output/RAKEN/excess_report.xlsx
output/AVTC/excess_report.xlsx
```

也可以直接用參數指定客戶，略過互動選擇：

```bash
python3 src/main.py --customer RAKEN
python3 src/main.py --customer AVTC
```

客戶模式會分別讀取：

```text
input/RAKEN/ -> output/RAKEN/excess_report.xlsx
input/AVTC/  -> output/AVTC/excess_report.xlsx
```

## 目前產出邏輯

- 讀取 `shortage` 的 Part No、Description、Planner、Buyer、LT、MOQ。
- 讀取 `shortage` 的 `Overshortage1`、`PO_REMAIN` 與第一欄 WO 外 demand。
- 若找不到 `Overshortage1`，會用 `OVER_SHORTAGE + HLD` 當替代。
- 從 BOM 的 G 欄判斷替代料：非 `*R*` 是主料，後續同一 BOM 區段連續 `*R*` 是替代料。
- 替代料不做跨機種/跨 BOM 區段的全域串聯；同一料號若在多個 BOM 區段出現，程式會挑選最適合目前 shortage 料號的一個區段。
- 已經被某一列 Excess 使用過的 shortage 料號，不會在後續列重複加總。
- 程式會先偵測所有 Excess 列需要的最大料號數，動態產出 `PartNo1` 到 `替代料N`。
- `OvershortageN`、`WO 外demandN`、`Open poN` 也會跟著 `替代料N` 自動增加。
- 逐料號的 `OvershortageN`、`WO 外demandN`、`Open poN` 會寫入從 shortage 算出的值。
- `Customer`、`MODEL` 會從 BOM 的 `成品料号` sheet 帶出；`MODEL` 使用選定 BOM 替代料區段對應的 C 欄主 Model。
- `MODELRemark/机种` 會從輸入 `BOM.xlsx` 的 `BOM` sheet B 欄帶出，並依輸出列的料號群組完整列出。
- Total 與 Excess 金額區會寫入 Excel 公式，讓使用者後續補 `Price（USD)` 或前版金額時可自動重算。
- 金額公式若受到缺少來源欄位影響，會先顯示空白；例如 `Price（USD)` 空白時，`Excess stockAmount`、`Excess POAmount`、`Excess TotalAMT` 會保持空白，避免把資料不足誤判成金額為 0。
- 數量與金額運算結果會四捨五入到整數輸出；`Price（USD)` 保留兩位小數供後續補單價。
- 輸出的 Excel 會套用 `Microsoft YaHei` 9 號字體，數值欄負數會以紅色顯示。

## Windows 執行檔

在 Windows 上可執行：

```bat
build_exe.bat
```

打包完成後會產生：

```text
release/BuyerExcessReports/BuyerExcessReports.exe
```

使用者收到整個 `BuyerExcessReports` 資料夾後，只要把檔案放進 `input/AVTC` 或 `input/RAKEN`，再雙擊 exe 並選擇客戶即可。

## 自訂路徑

也可以指定輸入與輸出；指定自訂路徑時不會強制要求選客戶：

```bash
python3 src/main.py --input-dir /path/to/input --output-file /path/to/excess_report.xlsx
```

如需關閉進度顯示，可加上：

```bash
python3 src/main.py --customer RAKEN --no-progress
```
