# assemble all vba file to one module
filenames = [
	'wrapper_function.vba',
	'updateSumDebt.vba',
	'importDeficitFile.vba',
	'qualResidYear.vba',
	'qualResidSeason.vba',
	'feeThisMonthV2.vba',
	'exportSheetAsXlsx.vba',
    'exportSheetAsXls97.vba',
	'genUpload.vba',
    'writeUploadCell.vba',
	'genAutoDeduct.vba',
    'genNotice.vba',
	'removeNoFee.vba',
    'removeNoFeeDeduct.vba',
    'renewDebt.vba',
	'ancmWordDoc.vba',
    'sendMail.vba',
    'res2csv.vba',
    'RunPythonInVenv.vba',
    'mergePdfPython.vba',
    'genFinance.vba'
]
# open in write mode will automatically clear file contents
with open('./whole_module.vba', 'w') as outfile:
    for fname in filenames:
        with open('./' + fname, encoding='utf-8') as infile:
            outfile.write(infile.read())

