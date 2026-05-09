# 翡翠流域大樓管理費收繳 Excel VBA and Python code

## 目錄結構與基本處理

管理費處理位於"自動化"目錄之下。VBA macro 則 ~~位於Macro子目錄下的Macro.xlam。~~ 放在
C:\Users\你的使用者名稱\AppData\Roaming\Microsoft\AddIns 之下。

VBA code 各依其不同功能獨立於各檔案。使用`python catAll.py`將所有更新集中到`whole_module.vba`檔案當中，再一次複製取代Macro.xlam原有的Module。

每個月的處理檔案各自獨立。主要檔案為"翡翠流域管理費收繳自動處理.xlsm"。此檔案的init為獨立的"WorkbookOpen.vba"。唯有此檔案不在Macro.xlam內。

## Excel使用說明:

### 下載Excel與Word檔案
* 翡翠流域管理費收繳自動處理.xlsm
* Macro.xlam
* 未銷帳代扣202604-4筆.xls
* 管理費自動扣繳通知.docm
* 公告範本.dotx

### 安裝增益集
將Macro.xlam放在
C:\Users\你的使用者名稱\AppData\Roaming\Microsoft\AddIns

在Excel中，選擇"檔案"->"選項"，選擇"增益集"->按下下方的"執行"
選擇"瀏覽"->開啟"Macro.xlam"

### 其他測試用檔案
可匯入的未銷帳資料:
未銷帳代扣202604-4筆.xls
此檔案為模擬元大銀行的未銷帳下載檔案，固定欄位，且格式為Excel97

## VBA Sub and Function 說明

* whole_module.vba 總檔案，與Macro.xlam內module的檔案吻合。
* wrapper_function.vba 內含許多Sub與Function。 這是公用程式與橋梁程式的集中檔案。
* importDeficitFile.vba 將元大銀行下載之未銷帳檔案，匯入"未銷帳"sheet。
* qualResidSeason.vba and qualResidYear.vba 通過月份與未銷帳，處理季繳與年繳資格問題。
* feeThisMonthV2.vba 計算當月應繳金額
* genUpload.vba 產生上傳資料(for 元大虛擬帳戶)
* genAutoDeduct.vba 產生自動代扣繳名冊
* sendMail.vba 寄出繳費通知email
* genFinance.vba 產生繳費財報頁(使用add-in menu)
* RunPythonInVenv.vba 呼叫split_pdf.py來執行pdf分頁(使用add-in menu)

## Python 輔助 code 說明

split_pdf.py 利用pypdf函式庫處理pdf檔案分割。 必須讀取listFile.csv以將pdf頁面對應到正確的戶名。 而listFile.csv則由上傳資料Excel檔， 利用uploadList2pdf所產生。 uploadList2pdf位於wrapper_function.vba當中。

merge_pdf.py 利用pypdf函式庫，將所有位於pdf子目錄下的pdf檔案合併。

## 相關連結

### 教學影片連結
* https://youtu.be/paewt2rhNPI
* https://youtu.be/jEjGHtR5fzo
* https://youtu.be/4L4PLj57KHc
* https://youtu.be/NxmJlCojGuA
* https://youtu.be/_Ifad1kQXnI
* https://youtu.be/00bn-fQcj_M
