# buyer-excess-reports

從 BOM、Item/Price 與含有 `shortage` 分頁的 Excel，自動產出可用版 `Excess` 報表。

程式會自動讀取檔名包含 `BOM` 的 Excel、檔名包含 `Item` 的 Excel，以及含有 `shortage` 分頁的 Excel。AVTC 使用 Item workbook 的「成本价」與 Price sheet，依序查主料成本價、主料 `Uprice`、同列替代料成本價、同列替代料 `Uprice`；RAKEN 會另讀檔名包含 `Price` 的 Excel，先查主料在 `Price` 與備援分頁的單價，再查同列替代料在這兩頁的單價。兩種客戶借用替代料價格時，單價格皆以紅字呈現。

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

首頁使用 `assets/app_background.png` 的木紋背景，配合較大的內容區、文字與按鈕；視窗會依螢幕空間等比縮放，最大寬度為 1440px。按 Start、選擇 AVTC 或 RAKEN 後，會先檢查對應 `input/<工廠>/` 內的來源檔案並產出 Excess 報表；完成後才會讀取本次 `excess_report.xlsx` 的 `MODEL` 欄，供使用者選擇要列入加單建議的機種。

MODEL 選擇頁支援搜尋、個別勾選及全選。按「確認選取」會產生 summary；第一張「加單建議（選）」沿用摘要頁的欄位與版型，依所選的 Excess 來源 MODEL 篩選摘要列。RAKEN 與 AVTC 的 Gating 都在「建議邏輯」中集中顯示每個 MODEL 依呆滯金額排序的 Top 1～Top 3、USD、TYPE、BOM MODEL、BOM 列、`EXTENDED_QT` 及計算方式，不再重複顯示前置 Top 資訊。Gating 不再放入加單建議文字；`Raken建议加单` 與 `AVTC建议加单` 欄則以「建議對BOM MODEL <MODEL>加單<數量>台」格式輸出完整建議。RAKEN 的 H、I 欄標題以 `Raken` 開頭；AVTC 維持 `AVTC` 標題。完整摘要頁統一名為 `加單建議（全）`；兩張摘要頁的完整 Gating 與 `庫存金額Top5` 均保留。Summary 工作表順序為 `加單建議（選）`、`加單建議（全）`、`供料明細`、`消耗建議`、`庫存金額Top5`。按「略過，不產出summary」時，本次不寫入 `excess_summary.xlsx`，其他報表照常產生；若輸出資料夾已有舊版 summary，程式會保留該檔、不會自動刪除。終端機互動模式也會提供 MODEL 搜尋與選取；非互動執行預設全選。

檔案確認畫面逐項顯示實心狀態圖示、實際檔名及缺漏或錯誤提示：綠色圓圈內白色勾號代表找到來源，紅色圓圈內白色叉號代表缺少必要來源或有錯誤，橘黃色圓圈內白色驚嘆號代表未提供選填的上一版。圖示由程式繪製，外觀不依賴 emoji 字型。右上角只顯示 `input/AVTC/` 或 `input/RAKEN/`。

BOM、Item、RAKEN 的 Price 依檔名關鍵字辨識；Shortage 會實際檢查 Excel 分頁名稱，不只看檔名。支援 `.xlsx`／`.xlsm`，忽略 `~$` 暫存檔。進入確認畫面先自動掃描一次；新增、移除或儲存來源檔後，請按主要確認按鈕右側的「重新掃描檔案」。每次點擊都會重新檢查，沒有定時自動掃描。掃描在背景執行，期間兩顆按鈕皆停用；若掃描期間來源改變，請再按一次「重新掃描檔案」。

必要檔案缺少時，主要按鈕顯示「請依照提示放入缺少檔案」，反灰且不能點擊或用鍵盤觸發；來源有重複候選或無法讀取時，也會提示並停用。補齊後按「重新掃描檔案」，檢查通過才改為「確認無誤，開始執行」，按下後會再次檢查再開始。兩顆按鈕整組置中。`next_time_excess_report` 為選填，未提供時顯示驚嘆號且仍可執行；若提供多份則須先排除重複。這一步檢查檔案與分頁是否可辨識，完整欄位及資料驗證仍在報表處理期間進行。

處理期間，介面會先顯示 Excess 報表處理進度（11 個階段）；Excess 報表完成後才開啟 MODEL 選擇頁，確認後再以 9 個階段產生 `excess_summary.xlsx`，包含讀取 Excess、BOM 配對、MODEL 彙總、加單建議、明細寫入與檔案儲存。略過或完成 Excess 報表時，不會產生本次 summary，但已產出的 Excess、next_time 與不完全替代料報表仍會保留。進度代表目前階段的完成數，各階段耗時不同。所有本次選擇要產生的報表都成功寫入後才會提示完成；按「確認」開啟對應 output 資料夾並關閉程式。若失敗，提示框會顯示實際檔案、處理階段及原因，按「確認」後關閉程式。處理中不接受關閉視窗，請等待結果。

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

AVTC 價格規則：

- 價格來自 Item workbook：「成本价」以 `Item` 查料號，取 `US Unit Cost`；Price sheet 以 `Part No.` 查料號，取 `Uprice`。
- 每列以報表的 `PartNo1` 為主料，即使該料號沒有出現在 shortage，仍優先查它的價格。依序查找：主料的成本價 → 主料的 `Uprice` → 同列所有替代料的成本價 → 同列所有替代料的 `Uprice`。
- 多個替代料時，同一分頁依 `替代料2`、`替代料3` 等欄位順序取第一個有效價格；所有替代料的成本價都優先於任何替代料的 `Uprice`。
- 找不到料號、價格空白或不是有效的有限數字時才繼續查下一步；`0` 是有效價格，不會回補。沒有「成本价」sheet 時使用 Price；只有有效成本資料而沒有 Price sheet 時也可產出，主料與替代料都缺價才留白。公式價格使用 Excel 儲存的計算結果，請先計算並儲存來源檔。
- 替代料僅限目前 BOM 嚴格使用情境規則認可、實際列在同一報表列的料號，包含未出現在 shortage 的同列替代料；不完全替代關係不納入借價。借用替代料價格時，`excess_report.xlsx` 與 `next_time_excess_report.xlsx` 的 `Price（USD)` 儲存格皆使用紅色字體，金額公式使用該單價計算；主料價格維持一般字色。執行紀錄會顯示替代料借價列數。

RAKEN 價格檔規則：

- 在 `input/RAKEN/` 這一層尋找檔名包含 `price` 的 `.xlsx`／`.xlsm`，不分大小寫，略過以 `~$` 開頭的暫存檔。符合的檔案必須只有一份；多份時會提示錯誤，不會依檔名日期自動挑選。
- 優先使用分頁名稱完全等於 `Price` 的 sheet，不分大小寫；名稱前後的空格不會忽略。備援固定使用排除該頁後，依 Excel 分頁從左到右的第一張，不依名稱排序、不要求日期或名稱包含 `RKN`。
- 若沒有 `Price` 分頁，僅使用最左邊的 sheet；若只有 `Price` 一頁，就只使用該頁。不會繼續搜尋其他分頁。
- 兩種表頭皆支援：以 `行标签` 查料號、取 `平均值项:US Unit Cost`，或以 `Part No.` 查料號、取 `USD`。表頭須在前 30 列內；選定分頁缺少必要表頭時會停止，並提示檔案與分頁。
- 每列以 `PartNo1` 為主料，依序查找：主料的 `Price` → 主料的備援頁 → 同列所有替代料的 `Price` → 同列所有替代料的備援頁。多個替代料時，同一分頁依 `替代料2`、`替代料3` 等欄位順序，取第一個有效價格；所有替代料的 `Price` 都優先於任何替代料的備援頁價格。
- 只有查不到料號、價格空白或不是有效的有限數字時，才繼續下一步。`0` 是有效價格，不會回補；主料與替代料在兩頁都缺價時才維持空白。公式價格沿用 Excel 儲存的計算結果，請先計算並儲存來源檔。
- 替代料僅限目前 BOM 嚴格使用情境規則認可、實際列在同一報表列的料號，包含未出現在 shortage 的同列替代料；被判定為不完全替代的關係不納入借價。借用替代料價格時，`excess_report.xlsx` 與 `next_time_excess_report.xlsx` 的 `Price（USD)` 儲存格皆使用紅色字體，金額公式會使用該單價計算；主料價格維持一般字色。
- RAKEN 的 `Price（USD)` 顯示四位小數，儲存值與金額公式仍使用原始單價精度。兩份輸出的 Excess sheet 皆建立兩個獨立、預設收合的欄位群組：`替代料2` 到最後一個替代料，以及 `Overshortage` 到 `Riskbuy` 的完整資訊區，包含總計與逐料號細項；按欄標上方的「＋」展開、「－」收合。沒有替代料時只建立資訊區群組，範圍隨實際欄位數量調整。第 1、2 列表頭保留自動換行，第 3 列起所有資料儲存格取消自動換行。

執行：

```bash
python3 src/main.py
```

互動終端機會先讓你選擇 `AVTC` 或 `RAKEN`，接著開啟 MODEL 選擇視窗；預設全選，可搜尋、個別勾選、全選，或略過 summary。若執行環境沒有可用的圖形桌面，會退回終端文字選單。執行時顯示 `tqdm` 進度條；非互動模式預設全選 MODEL。

輸出：

```text
output/RAKEN/excess_report.xlsx
output/RAKEN/next_time_excess_report.xlsx
output/RAKEN/excess_summary.xlsx（未略過 summary 時產生）
output/RAKEN/incomplete_alternates.xlsx
output/AVTC/excess_report.xlsx
output/AVTC/next_time_excess_report.xlsx
output/AVTC/excess_summary.xlsx（未略過 summary 時產生）
output/AVTC/incomplete_alternates.xlsx
```

也可以直接用參數指定客戶，略過互動選擇：

```bash
python3 src/main.py --customer RAKEN
python3 src/main.py --customer AVTC
```

客戶模式會分別讀取與輸出：

```text
input/RAKEN/ -> output/RAKEN/excess_report.xlsx + output/RAKEN/next_time_excess_report.xlsx + output/RAKEN/excess_summary.xlsx（未略過時） + output/RAKEN/incomplete_alternates.xlsx
input/AVTC/  -> output/AVTC/excess_report.xlsx  + output/AVTC/next_time_excess_report.xlsx  + output/AVTC/excess_summary.xlsx（未略過時） + output/AVTC/incomplete_alternates.xlsx
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
- AVTC 與 RAKEN 的 `Customer` 基礎來源都是 Item sheet 的 `Mps Customer`。RAKEN 若有上一版 `next_time_excess_report.xlsx`，會依料號配對直接承接上一版 Customer 值；未配對的列留白。沒有上一版來源時使用本期 Item 值。
- `MODEL` 會優先從 BOM 的 `成品料号` sheet 帶出；若找不到，改用 Item sheet 的 `Mps Model`。
- `MODELRemark/机种` 會從輸入 BOM 的 `BOM` sheet B 欄帶出，並依輸出列的料號群組完整列出。
- RAKEN 的 `Customer`、`MODEL`、`MODELRemark/机种` 會放在 `GSD` 欄後方；AVTC 會將這三欄放在 `reason&action` 後方，對齊 1006 版 AVTC Excess 版型。
- `Price（USD)` 來源依客戶 profile 讀取：AVTC 與 RAKEN 分別依上述主料、替代料及分頁優先順序查價，借用替代料單價時皆標紅。同一分頁遇到同料號多筆有效價格時，仍優先取 `Allocation` 最大者，若相同再取 `Start Date` 較新的，仍相同則取分頁中較前面的第一筆。
- Total 與 Excess 金額區會寫入 Excel 公式，讓使用者後續補前版金額時可自動重算。
- 程式每次會多輸出 `next_time_excess_report.xlsx`。若下次執行時 input 資料夾內有檔名包含 `next_time` 的 Excel，程式會將它視為上一版來源，用第一個 `Excess TotalAMT` 區間帶入 `Previous Excess Total AMT`，並輸出 `Improve` 公式；若找不到上一版來源，終端機會顯示 `找不到上一版來源，不輸出improve資訊`，且 `Improve` 維持空白。
- RAKEN 會依 `PartNo1` 配對上一版；若主料號順序改變，再以完整替代料群組配對。`MODEL`、`Customer`、第一段可手動輸入的 `PO外備料 QTY`、`reason&action`、`Vendor` 會一起承接到兩份輸出，包含使用者刻意清空的儲存格。若提供上一版來源但找不到配對列，Customer 留白；沒有上一版來源時 Customer 使用本期 Item 的 `Mps Customer`。其他新料號資料維持本期來源。上一版有重複主料號或群組時會停止並提示，避免不明確的資料配對。
- AVTC 以相同料號／完整替代料群組規則，將上一版手動更新的 `reason&action`、`Vendor` 承接到 `excess_report.xlsx` 與 `next_time_excess_report.xlsx`，保留多行文字及刻意清空的儲存格。新料號沒有上一版資料時，這兩欄維持空白；重複主料號或群組造成不明確配對時會停止並提示。AVTC 的 MODEL、Customer 維持本期來源。
- Improve 維持「本期第一段 Excess Total AMT－上一版第一段 Excess Total AMT」。上一版金額依上一版資料與單價計算，不會用本期新成本重算上一版；價格變動也可能影響 Improve。
- 上一版支援舊版與新版備料欄名，並自動辨識實際表頭列；輸出維持兩列表頭，不複製操作說明列。人工欄位若使用公式，會承接 Excel 儲存的計算結果；匯入前須先在 Excel 重算並儲存。
- 報表中的金額等計算欄使用 Excel 2016 可用的傳統公式，相容目標為 Excel 2016、2021、2024 與 Microsoft 365。RAKEN 的 `Customer客户` 欄是一般儲存值，不使用公式查表；上一版 Customer 儲存格若是舊版公式，程式承接 Excel 儲存的計算結果，舊值為 `error` 或沒有快取結果時留白。
- RAKEN 若上一版有 `Customer` 工作表，兩份輸出會保留該表作為舊資料存檔，但不再用它填入 Customer 欄；不會產生 `_CustomerLookup` 輔助分頁。AVTC 承接 `reason&action`、`Vendor`，Customer 使用本期 Item 資料。
- 數量與金額運算結果會四捨五入到整數輸出；`Price（USD)` 在 RAKEN 顯示四位小數、AVTC 顯示兩位小數，金額公式皆使用原始單價精度。
- AVTC 預設隱藏 `Overshortage1` 到 `OvershortageN`、各時間區間的 `WO 外demand1` 到 `WO 外demandN`、以及 `Open po1` 到 `Open poN` 細項欄，保留總欄位顯示。RAKEN 則使用上述兩個可展開的欄位群組。
- RAKEN 每段 Excess 欄位順序為 `Excess POQty`、`Excess POAmount`、`Excess stockQTY`、`Excess stockAmount`、`Excess TotalAMT`；`Excess TotalAMT` 會加總 `Excess POAmount`、`Excess stockAmount`、`Excess  PO外備料AMT`。`reason&action`、`Vendor` 整欄會依此順序放在第一段 Excess 之後、第二段備料欄位之前。
- AVTC 每段 Excess 欄位順序為 `Excess stockQTY`、`Excess stockAmount`、`Excess POQty`、`Excess POAmount`、`Excess TotalAMT`；`Excess TotalAMT` 會加總 `Excess POAmount` 與 `Excess stockAmount`，且每段 Excess 標題會使用不同底色。
- AVTC 與 RAKEN 可輸出 `excess_summary.xlsx`，工作表依序為 `加單建議（選）`、`加單建議（全）`、`供料明細`、`消耗建議`、`庫存金額Top5`。兩張加單建議頁都保留完整欄位；RAKEN 會在 A 欄後顯示 `Customer 客戶`，A 欄標題依流程顯示「AVTC」或「Raken」。RAKEN 的 H、I 欄標題分別以 `Raken` 開頭，AVTC 維持 `AVTC` 標題。程式會先產出本次 `excess_report.xlsx`，再讀取該報表 `MODEL` 欄的唯一非空值，讓使用者搜尋並選取要列入 Summary 的 MODEL。`加單建議（選）` 依使用者選取的 Excess 報表 MODEL 篩選列；`加單建議（全）` 維持完整資料，不受選取範圍影響。RAKEN 與 AVTC 的 Gating 都在「建議邏輯」中集中顯示每個 MODEL 依呆滯金額排序的 Top 1～Top 3、USD、TYPE、BOM MODEL、BOM 列、`EXTENDED_QT` 及換算方式，不再重複顯示前置 Top 資訊，不再於 Gating 顯示加單建議；「Raken建议加单」或「AVTC建议加单」欄與完整摘要頁一致，格式為「建議對BOM MODEL <MODEL>加單<數量>台」。已選但沒有 Excess 摘要資料的 MODEL 會列出備註；`庫存金額Top5` 維持完整資料，不受選取範圍影響。按略過會跳過本次 summary 寫入，但已產出的 Excess、next_time 與不完全替代料報表仍會保留。內容依最後一期的 Excess 欄位彙總至加單建議頁，明細列於 `庫存金額Top5`；摘要與 `庫存金額Top5` 依本次 `excess_report` 的 `MODEL` 欄分組，不再用 BOM B 欄決定摘要機種。BOM B 欄仍只用於找出 BOM 用量、`TOP_ASSEMBLY_ITEM`、`EXTENDED_QT` 與供料計算；因此兩者不同時，Summary 會以 Excess 報表 MODEL 歸類，但保留 BOM MODEL 供追溯。沒有可用 Excess MODEL 的列才歸入「BOM 未配對」，每個 MODEL 各自依 `Excess stockAmount` 取前五筆。庫存金額Top5仍列五名，並移除 MODELRemark 與 Desc。建議依每筆 Excess 庫存料號及替代料配對 BOM 的 `MODEL`、`TOP_ASSEMBLY_ITEM`、`EXTENDED_QT` 換算，不讀 Shortage PP TTL。每筆庫存只分配給一個可行的 BOM 用料情境，優先選每台用量較高者；用量相同時依 BOM 列序、MODEL、成品料號決定。每筆建議獨立以該筆 Excess stock QTY ÷ 對應 BOM EXTENDED_QT 計算所需台數（向上取整），加單目標仍以 BOM MODEL 為準；建議對象依配對結果的 BOM MODEL 決定。消耗金額只列該筆 Excess 庫存金額，不合併或共用不同庫存項目的加單台數。相同 MODEL、成品料號、相同替代料群的 BOM 用量會先合併；摘要可使用 BOM 單列明列的替代料及該列用量，不要求替代料涵蓋主料在其他成品的所有使用情境；一般 Excess 報表的替代料規則不變。其他 BOM 料件由一般供料流程供應。摘要數值寫成固定值，不使用 MSI 範本公式或 Excess 報表的公式快取。RAKEN 與 AVTC 的加單建議均使用 BOM 的 `MODEL` 欄作為目標；`TOP_ASSEMBLY_ITEM` 用於識別 BOM 用料情境及追溯用量。加單建議不讀 Shortage／PP 需求；Shortage 僅供產生 Excess 報表本身使用。RAKEN 摘要另依 Excess 欄位口徑計算：`Excess stockQTY` 扣除 `others demand`，`Excess TotalAMT` 納入 `Excess  PO外備料AMT`；該備料金額依上一版承接的 `PO外備料 QTY` 計算。若來源沒有可辨識的 BOM MODEL、用量或成品料號，摘要及 Top 5 仍會產生，但加單欄會說明原因。
- `excess_summary.xlsx` 的一般候選料號還會套用內建的 Item TYPE 規則：BOM 實際配對的 `matched_part` 必須在 `DP`、`Film`、`IC`、`Interconnection`、`LGP`、`Light Bar`、`Metal`、`Module`、`OC`、`PCB`、`PCBA`、`PD2-1`、`PD2-2`、`PD2-3`、`PD2-O`、`Plastic`、`Power`、`Remote`、`SMT-I-1`、`SMT-O-1`、`SOC`、`SPK` 中，才會進入 Top 5／Gating Top 3／加單建議；以 `3` 開頭且成功配對 BOM D 欄 `L ASSEMBLY_ITEM` 的半成品不受這份 TYPE 名單限制。Gating 的「建議邏輯」內會顯示每個配對料號的 TYPE；半成品若 TYPE 空白，會標示為「半成品免 TYPE 篩選」。AVTC 從 `Item` 讀取 TYPE，RAKEN 從 `Item information` 讀取 TYPE；不依賴 `~/Downloads/Item_fix.xlsx`。替代料群組只檢查 `matched_part`，不要求群組內所有替代料都符合。
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

使用者收到整個 `BuyerExcessReports` 資料夾後，只要把檔案放進 `input/AVTC` 或 `input/RAKEN`，再雙擊 exe、按 Start 並選擇工廠；程式會先產出 Excess 報表，再從本次報表的 `MODEL` 欄開啟選擇介面，讓使用者搜尋、勾選 MODEL 或略過 summary。Windows 版本使用 `--windowed` 打包，執行期間不會開啟終端機；圖片、Tkinter 與 Pillow 都會一起打包，使用者不需要安裝 Python，也不需要原本 Downloads 的圖片。

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

Excess summary 的最新顯示規則：RAKEN 與 AVTC 的 Gating 改在「建議邏輯」內集中顯示每個 MODEL 的 Top 1～Top 3、庫存金額、TYPE、BOM MODEL、TOP_ASSEMBLY_ITEM、BOM 列、一般料號的 `EXTENDED_QT` 或半成品的每台用量，以及對應計算方式，不再在前方重複顯示 Top 與 BOM 對應資訊。新的加單消耗與供料明細不塞入 Gating。summary 以剛產生的 Excess 報表 `MODEL` 欄分類；BOM B 欄只作為 BOM 用量與供料配對資訊，不再覆蓋 Summary 的機種歸屬。摘要頁名稱為 `加單建議（選）` 與 `加單建議（全）`，並依序提供 `供料明細`、`消耗建議`、`庫存金額Top5`。各工作表標題已依欄位用途分組套用底色；加單建議頁的多筆建議會以空白行分隔，需額外採購摘要會逐項換行並在最後標示其餘筆數。摘要頁新增 `消耗率`、建議料號可消耗金額、所有相關料號原有 Excess USD（加單前）、所有相關料號可消耗金額、加單後剩餘 Excess、庫存數量是否可滿足加單及需額外採購欄位；`所有相關料號原有 Excess USD（加單前）` 為消耗率分母，並等於所有相關料號可消耗 USD 加上加單後剩餘 Excess USD。`Raken建议加单` 與 `AVTC建议加单` 欄以建議編號對應每筆建議。

`excess_summary.xlsx` 另新增 `消耗建議` 與 `供料明細` 工作表。`消耗建議` 每筆建議一列，列出 BOM MODEL、TOP_ASSEMBLY_ITEM、加單台數、消耗率、所有相關料號原有 Excess USD（加單前）、可消耗金額與剩餘金額；不足摘要以「共 N 筆料號／替代料群不足」呈現，實際採購數量仍以 `供料明細` 的「需額外採購數量」欄位為準。`供料明細` 逐一列出必要料號／替代料群的每台用量、加單需求、現有 Excess stock QTY、現有 Excess USD、OPO、可消耗數量、需額外採購數量及剩餘 Excess。每筆建議獨立計算，OPO 僅作為參考，不納入現有 Excess 消耗率；RAKEN 沿用扣除 `others demand` 後的 Excess stock QTY。

若候選料號以 `3` 開頭，Summary 會將其視為半成品，改以 BOM D 欄 `L ASSEMBLY_ITEM` 配對，而不是以 H 欄 `COMPONENT_ITEM` 配對；只要成功配對到 D 欄，即使 Item TYPE 空白或不在一般 TYPE 名單內，也可進入後續篩選。未配對到 D 欄的料號仍不會因為以 `3` 開頭而放行。供料計算會先扣除該半成品的 Excess 庫存；半成品不足時，才依 BOM 階層展開下階 H 欄物料，逐項扣除下階 Excess，剩餘需求列為須額外採購。`供料明細` 會以「半成品」、「下階半成品」或「下階物料」標示供料角色，並避免同一庫存列被重複消耗。

`加單建議（選）` sheet 的 A 欄後新增 `Customer 客戶` 欄位，RAKEN 的 `加單建議（全）` 也在 A 欄後顯示相同的 Customer 資訊；Summary 會在 Excess report 寫入後，直接讀取該報表的 `MODEL` 與 `Customer` 欄位。Customer 值會移除 `/` 後的 `#N`、`#A`、`#N/A`、`N/A` 等無效片段，例如 `ASUS/#N/#A` 顯示為 `ASUS`。
