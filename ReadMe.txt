vba_backup05
2026/04/18
1. change all macro reference from *.xlsm itself to ..\macro\macro.xlam, as an add-in
This is for solving other newly created Excel file, would always reference to older main Excel file.
Note WorkbookOpen.vba still sit with *.xlsm, as it is the initial process. It only write to log and clear checkmarks, nothing important.