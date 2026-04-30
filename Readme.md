# 翡翠流域大樓管理費收繳 Excel VBA and Python code

## 目錄結構與基本處理

管理費處理位於"自動化"目錄之下。VBA macro 則位於Macro子目錄下的Macro.xlam。

VBA code 各依其不同功能獨立於各檔案。使用`python catAll.py`將所有更新集中到`whole_module.vba`檔案當中，再一次複製取代Macro.xlam原有的Module。

每個月的處理檔案各自獨立。主要檔案為"翡翠流域管理費收繳自動處理.xlsm"。此檔案的init為獨立的"WorkbookOpen.vba"。唯有此檔案不在Macro.xlam內。

## VBA Sub and Function 說明

* whole_module.vba 總檔案，與Macro.xlam內module的檔案吻合。
* wrapper_function.vba 內含許多Sub與Function。 這是公用程式與橋梁程式的集中檔案。
* importDeficitFile.vba 將元大銀行下載之未銷帳檔案，匯入"未銷帳"sheet。
* qualResidSeason.vba and qualResidYear.vba 通過月份與未銷帳，處理季繳與年繳資格問題。
* feeThisMonthV2.vba 計算當月應繳金額
* genUpload.vba 產生上傳資料(for 元大虛擬帳戶)
* genAutoDeduct.vba 產生自動代扣繳名冊
* sendMail.vba 寄出繳費通知email

## Python 輔助 code 說明

split_pdf.py 利用pypdf函式庫處理pdf檔案分割。 必須讀取listFile.csv以將pdf頁面對應到正確的戶名。 而listFile.csv則由上傳資料Excel檔， 利用uploadList2pdf所產生。 uploadList2pdf位於wrapper_function.vba當中。

merge_pdf.py 利用pypdf函式庫，將所有位於pdf子目錄下的pdf檔案合併。