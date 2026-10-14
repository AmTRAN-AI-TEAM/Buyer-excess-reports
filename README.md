# buyer-excess-reports

從 BOM、Item/Price 與含有 `shortage` 分頁的 Excel，自動產出可用版 `Excess` 報表。

程式會自動讀取檔名包含 `BOM` 的 Excel、檔名包含 `Item` 的 Excel，以及含有 `shortage` 分頁的 Excel。AVTC 價格來自 Item workbook 的 Price sheet `Uprice` 欄；RAKEN 會另讀檔名包含 `Price` 的 Excel，價格來自 `USD` 欄。

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
input/RAKEN/Item.xlsx
input/RAKEN/Price.xlsx
input/RAKEN/你的shortage檔.xlsx

input/AVTC/BOM.xlsx
input/AVTC/Item.xlsx
input/AVTC/你的shortage檔.xlsx
```

BOM 檔名只要包含 `BOM` 字樣即可；Item 檔名只要包含 `Item` 字樣即可；RAKEN 價格檔檔名只要包含 `Price` 字樣即可。shortage 檔內只要有分頁名稱包含 `shortage`，程式就會使用該分頁。輸入檔不需要有 `Excess` sheet。

執行：

```bash
python3 src/main.py
```

程式會先讓你選擇 `AVTC` 或 `RAKEN`，再讀取對應的 `input/<客戶>/`，並在執行時顯示 `tqdm` 進度條。

輸出：

```text
output/RAKEN/excess_report.xlsx
output/RAKEN/next_time_excess_report.xlsx
output/RAKEN/incomplete_alternates.xlsx
output/AVTC/excess_report.xlsx
output/AVTC/next_time_excess_report.xlsx
output/AVTC/incomplete_alternates.xlsx
```

也可以直接用參數指定客戶，略過互動選擇：

```bash
python3 src/main.py --customer RAKEN
python3 src/main.py --customer AVTC
```

客戶模式會分別讀取與輸出：

```text
input/RAKEN/ -> output/RAKEN/excess_report.xlsx + output/RAKEN/next_time_excess_report.xlsx + output/RAKEN/incomplete_alternates.xlsx
input/AVTC/  -> output/AVTC/excess_report.xlsx  + output/AVTC/next_time_excess_report.xlsx  + output/AVTC/incomplete_alternates.xlsx
```

## 目前產出邏輯

- 讀取 `shortage` 的 Part No（`COMPONENT_ITEM` 也視為 Part No）、Description、Planner、Buyer、LT、MOQ。
- 讀取 `shortage` 的 `Overshortage1`、`PO_REMAIN`，以及 Part No 前方所有 `WO外demand` / `工单外...demand` 欄位。
- 多個 `WO外demand` 欄位會依上一列表頭作為時間區間，例如 `cover Jan'28`、`cover 'Dec`、`cover 'Jun`，並在輸出中展開對應的 WO demand 與 Excess 區塊。
- 若找不到 `Overshortage1`，會用 `OVER_SHORTAGE + HLD` 當替代。
- 從 BOM 的 G 欄判斷候選替代料：非 `*R*` 是主料，後續同一 BOM 區段連續 `*R*` 是候選替代料。
- 替代料採嚴格 BOM 使用情境規則：同一替代料必須出現在主料所有 BOM 使用情境中，才會保留在同一組；若主料在其他 BOM 使用情境沒有該替代料，該替代關係會被拆開。
- 被嚴格規則拆開的候選替代關係會另外輸出 `incomplete_alternates.xlsx`，欄位包含 Item、Model、Part No、Spec；每個不完全替代群組內，同一料號只列一行，Model 會彙總該料號在 BOM B 欄出現過的所有機種，Spec 來源為 BOM 的 `COMPONENT_DESCRIPTION` 第一個非空值。
- 替代料不做跨機種/跨 BOM 區段的全域串聯；同一料號若在多個合格 BOM 區段出現，程式會挑選最適合目前 shortage 料號的一個區段。
- 已經被某一列 Excess 使用過的 shortage 料號，不會在後續列重複加總。
- 程式會先偵測所有 Excess 列需要的最大料號數，動態產出 `PartNo1` 到 `替代料N`。
- `OvershortageN`、各時間區間的 `WO 外demandN`、`Open poN` 也會跟著 `替代料N` 自動增加。
- 逐料號的 `OvershortageN`、`WO 外demandN`、`Open poN` 會寫入從 shortage 算出的值。
- `Customer` 來源改為 Item sheet 的 `Mps Customer`。
- `MODEL` 會優先從 BOM 的 `成品料号` sheet 帶出；若找不到，改用 Item sheet 的 `Mps Model`。
- `MODELRemark/机种` 會從輸入 BOM 的 `BOM` sheet B 欄帶出，並依輸出列的料號群組完整列出。
- RAKEN 的 `Customer`、`MODEL`、`MODELRemark/机种` 會放在 `GSD` 欄後方；AVTC 會將這三欄放在 `reason&action` 後方，對齊 1006 版 AVTC Excess 版型。
- `Price（USD)` 來源依客戶 profile 讀取：AVTC 使用 `Uprice`，RAKEN 使用獨立 Price 檔的 `USD`；同料號多筆價格時，優先取 `Allocation` 最大者，若相同再取 `Start Date` 較新的，仍相同則取檔案中較前面的第一筆。
- Total 與 Excess 金額區會寫入 Excel 公式，讓使用者後續補前版金額時可自動重算。
- 程式每次會多輸出 `next_time_excess_report.xlsx`。若下次執行時 input 資料夾內有檔名包含 `next_time` 的 Excel，程式會將它視為上一版來源，用第一個 `Excess TotalAMT` 區間帶入 `Previous Excess Total AMT`，並輸出 `Improve` 公式；若找不到上一版來源，終端機會顯示 `找不到上一版來源，不輸出improve資訊`，且 `Improve` 維持空白。
- 數量與金額運算結果會四捨五入到整數輸出；`Price（USD)` 保留兩位小數。
- 輸出會預設隱藏 `Overshortage1` 到 `OvershortageN`、各時間區間的 `WO 外demand1` 到 `WO 外demandN`、以及 `Open po1` 到 `Open poN` 細項欄，保留總欄位顯示。
- RAKEN 每段 Excess 欄位順序為 `Excess POQty`、`Excess POAmount`、`Excess stockQTY`、`Excess stockAmount`、`Excess TotalAMT`；`Excess TotalAMT` 會加總 `Excess POAmount`、`Excess stockAmount`、`PO外備料AMT`。
- AVTC 每段 Excess 欄位順序為 `Excess stockQTY`、`Excess stockAmount`、`Excess POQty`、`Excess POAmount`、`Excess TotalAMT`；`Excess TotalAMT` 會加總 `Excess POAmount` 與 `Excess stockAmount`，且每段 Excess 標題會使用不同底色。
- RAKEN 輸出會在第一個 Excess 區塊前插入 `others demand`，並在每個 Excess 區塊前插入 `PO外備料`、`PO外備料QTY`、`PO外備料AMT`。其中 `others demand` 與 `PO外備料` 由使用者填寫，`PO外備料QTY` 與 `PO外備料AMT` 由公式計算。
- AVTC 不輸出 `others demand`、`PO外備料`、`PO外備料QTY`、`PO外備料AMT`，並會在 `Riskbuy` 與第一段 Excess 中間保留兩個空白欄。
- RAKEN 每段 `Excess stockQTY` 會扣除該時間區間的 `WO外demand`，並再扣除 `others demand`；AVTC 每段 `Excess stockQTY` 只扣除該時間區間的 `WO外demand`。
- 輸出的 Excel 會套用 `Microsoft YaHei` 9 號字體，標題會依資料區域套用不同底色，數值欄負數會以紅色顯示。

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
