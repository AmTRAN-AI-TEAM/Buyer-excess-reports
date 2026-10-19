# buyer-excess-reports

從 BOM、Item/Price 與含有 `shortage` 分頁的 Excel，自動產出可用版 `Excess` 報表。

程式會自動讀取檔名包含 `BOM` 的 Excel、檔名包含 `Item` 的 Excel，以及含有 `shortage` 分頁的 Excel。AVTC 價格優先來自 Item workbook 的「成本价」sheet（`Item`、`US Unit Cost`），逐料號缺價時才使用原本 Price sheet 的 `Uprice` 欄；RAKEN 會另讀檔名包含 `Price` 的 Excel，優先讀取名稱完全符合 `Price` 的分頁，逐料號缺價時使用其餘分頁中最左邊的一張補上。

## 資料夾

```text
buyer-excess-reports/
  input/
    AVTC/
    RAKEN/
  output/
    AVTC/
    RAKEN/
  assets/
    app_background.png
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

桌面介面可執行：

```bash
python3 generate_buyer_excess_report.py
```

首頁使用 `assets/app_background.png` 的木紋背景，配合較大的內容區、文字與按鈕；視窗會依螢幕空間等比縮放，最大寬度為 1440px。按 Start、選擇 AVTC 或 RAKEN 後，程式會自動檢查對應 `input/<工廠>/` 內的來源檔案。

檔案確認畫面逐項顯示實心狀態圖示、實際檔名及缺漏或錯誤提示：綠色圓圈內白色勾號代表找到來源，紅色圓圈內白色叉號代表缺少必要來源或有錯誤，橘黃色圓圈內白色驚嘆號代表未提供選填的上一版。圖示由程式繪製，外觀不依賴 emoji 字型。右上角只顯示 `input/AVTC/` 或 `input/RAKEN/`。

BOM、Item、RAKEN 的 Price 依檔名關鍵字辨識；Shortage 會實際檢查 Excel 分頁名稱，不只看檔名。支援 `.xlsx`／`.xlsm`，忽略 `~$` 暫存檔。進入確認畫面先自動掃描一次；新增、移除或儲存來源檔後，請按主要確認按鈕右側的「重新掃描檔案」。每次點擊都會重新檢查，沒有定時自動掃描。掃描在背景執行，期間兩顆按鈕皆停用；若掃描期間來源改變，請再按一次「重新掃描檔案」。

必要檔案缺少時，主要按鈕顯示「請依照提示放入缺少檔案」，反灰且不能點擊或用鍵盤觸發；來源有重複候選或無法讀取時，也會提示並停用。補齊後按「重新掃描檔案」，檢查通過才改為「確認無誤，開始執行」，按下後會再次檢查再開始。兩顆按鈕整組置中。`next_time_excess_report` 為選填，未提供時顯示驚嘆號且仍可執行；若提供多份則須先排除重複。這一步檢查檔案與分頁是否可辨識，完整欄位及資料驗證仍在報表處理期間進行。

處理期間，介面會顯示即時執行紀錄與 11 個階段的進度。進度代表已完成的階段數，各階段耗時不同。所有報表都成功寫入後才會提示完成；按「確認」開啟對應 output 資料夾並關閉程式。若失敗，提示框會顯示實際檔案、處理階段及原因，按「確認」後關閉程式。處理中不接受關閉視窗，請等待結果。

執行紀錄以 20px 為基準顯示（1200px 寬視窗），隨視窗等比縮放，並增加行距與紀錄區高度。處理畫面顯示目前階段、已完成階段數、百分比與即時紀錄，不再顯示預估剩餘時間，也不再讀寫 `.app_state/run_timings.json`。

命令列仍可使用下方的 `src/main.py` 與路徑參數。

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

RAKEN 價格檔規則：

- 在 `input/RAKEN/` 這一層尋找檔名包含 `price` 的 `.xlsx`／`.xlsm`，不分大小寫，略過以 `~$` 開頭的暫存檔。符合的檔案必須只有一份；多份時會提示錯誤，不會依檔名日期自動挑選。
- 優先使用分頁名稱完全等於 `Price` 的 sheet，不分大小寫；名稱前後的空格不會忽略。備援固定使用排除該頁後，依 Excel 分頁從左到右的第一張，不依名稱排序、不要求日期或名稱包含 `RKN`。
- 若沒有 `Price` 分頁，僅使用最左邊的 sheet；若只有 `Price` 一頁，就只使用該頁。不會繼續搜尋其他分頁。
- 兩種表頭皆支援：以 `行标签` 查料號、取 `平均值项:US Unit Cost`，或以 `Part No.` 查料號、取 `USD`。表頭須在前 30 列內；選定分頁缺少必要表頭時會停止，並提示檔案與分頁。
- 逐料號在優先分頁查不到、價格空白或不是有效的有限數字時，才使用備援頁。`0` 是有效價格，不會回補；兩頁都缺價時維持空白。公式價格沿用 Excel 儲存的計算結果，請先計算並儲存來源檔。

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
- `Price（USD)` 來源依客戶 profile 讀取：AVTC 優先以 Item workbook「成本价」sheet 的 `Item` 查料號，取 `US Unit Cost`；料號不存在、價格空白或不是有效數字時，回退至原本 Price sheet 的 `Uprice`。成本為 0 視為有效價格，不會回退。若沒有「成本价」sheet，沿用原本價格來源；若只有有效成本資料而沒有 Price sheet，仍可產出報表，沒有任何可用價格的料號維持價格空白。RAKEN 依上述分頁優先順序使用獨立 Price 檔的價格。同一分頁遇到同料號多筆有效價格時，仍優先取 `Allocation` 最大者，若相同再取 `Start Date` 較新的，仍相同則取分頁中較前面的第一筆；RAKEN 的 `Price` 有有效價格時，優先於備援頁的任何價格。
- Total 與 Excess 金額區會寫入 Excel 公式，讓使用者後續補前版金額時可自動重算。
- 程式每次會多輸出 `next_time_excess_report.xlsx`。若下次執行時 input 資料夾內有檔名包含 `next_time` 的 Excel，程式會將它視為上一版來源，用第一個 `Excess TotalAMT` 區間帶入 `Previous Excess Total AMT`，並輸出 `Improve` 公式；若找不到上一版來源，終端機會顯示 `找不到上一版來源，不輸出improve資訊`，且 `Improve` 維持空白。
- RAKEN 會依 `PartNo1` 配對上一版；若主料號順序改變，再以完整替代料群組配對。`MODEL`、第一段可手動輸入的 `PO外備料 QTY`、`reason&action`、`Vendor` 會一起承接到兩份輸出，包含使用者刻意清空的儲存格；新料號維持本期原有資料。上一版有重複主料號或重複群組時會停止並提示，避免不明確的資料配對。
- AVTC 以相同料號／完整替代料群組規則，將上一版手動更新的 `reason&action`、`Vendor` 承接到 `excess_report.xlsx` 與 `next_time_excess_report.xlsx`，保留多行文字及刻意清空的儲存格。新料號沒有上一版資料時，這兩欄維持空白；重複主料號或群組造成不明確配對時會停止並提示。AVTC 的 MODEL、Customer 維持本期來源。
- Improve 維持「本期第一段 Excess Total AMT－上一版第一段 Excess Total AMT」。上一版金額依上一版資料與單價計算，不會用本期新成本重算上一版；價格變動也可能影響 Improve。
- 上一版支援舊版與新版備料欄名，並自動辨識實際表頭列；輸出維持兩列表頭，不複製操作說明列。人工欄位若使用公式，會承接 Excel 儲存的計算結果；匯入前須先在 Excel 重算並儲存。
- RAKEN 會先承接 `MODEL`，再用上一版 `Customer` sheet 的 `MODEL`、`Customer` 欄查表，取代本期 `Customer客户`。先完整比對 MODEL；找不到時，以 `/` 拆分多機種，全部查到才將客戶去重並以 `/` 合併。MODEL 空白、查不到、部分查不到或同機種有衝突對應時，保留本期 Item 來源的 Customer，終端機會顯示筆數及部分範例。
- RAKEN 兩份輸出都會保留 Customer 對照表資料；沒有上一版 Customer sheet 時，建立只有 `MODEL`、`Customer` 表頭的空白 Customer sheet，供後續維護。AVTC 承接 `reason&action`、`Vendor`，Customer 使用本期 Item 資料。
- 數量與金額運算結果會四捨五入到整數輸出；`Price（USD)` 保留兩位小數。
- 輸出會預設隱藏 `Overshortage1` 到 `OvershortageN`、各時間區間的 `WO 外demand1` 到 `WO 外demandN`、以及 `Open po1` 到 `Open poN` 細項欄，保留總欄位顯示。
- RAKEN 每段 Excess 欄位順序為 `Excess POQty`、`Excess POAmount`、`Excess stockQTY`、`Excess stockAmount`、`Excess TotalAMT`；`Excess TotalAMT` 會加總 `Excess POAmount`、`Excess stockAmount`、`Excess  PO外備料AMT`。`reason&action`、`Vendor` 整欄會依此順序放在第一段 Excess 之後、第二段備料欄位之前。
- AVTC 每段 Excess 欄位順序為 `Excess stockQTY`、`Excess stockAmount`、`Excess POQty`、`Excess POAmount`、`Excess TotalAMT`；`Excess TotalAMT` 會加總 `Excess POAmount` 與 `Excess stockAmount`，且每段 Excess 標題會使用不同底色。
- RAKEN 輸出會在第一個 Excess 區塊前插入 `others demand`，並在每個 Excess 區塊前插入 `PO外備料 QTY`、`Excess PO外備料QTY`、`Excess  PO外備料AMT`。其中 `others demand` 與第一段 `PO外備料 QTY` 由使用者填寫；後續各段 `PO外備料 QTY` 以公式引用第一段，請只修改第一段。各段 Excess 備料數量與金額仍依該段需求重新計算。
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

使用者收到整個 `BuyerExcessReports` 資料夾後，只要把檔案放進 `input/AVTC` 或 `input/RAKEN`，再雙擊 exe、按 Start 並選擇工廠即可。Windows 版本使用 `--windowed` 打包，執行期間不會開啟終端機；圖片、Tkinter 與 Pillow 都會一起打包，使用者不需要安裝 Python，也不需要原本 Downloads 的圖片。

請在 Windows 執行 `build_exe.bat`，產生可供 Windows 使用的執行檔。

## 自訂路徑

也可以指定輸入與輸出；指定自訂路徑時不會強制要求選客戶：

```bash
python3 src/main.py --input-dir /path/to/input --output-file /path/to/excess_report.xlsx
```

如需關閉進度顯示，可加上：

```bash
python3 src/main.py --customer RAKEN --no-progress
```
