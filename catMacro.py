# assemble all vba file to one module
filenames = [
	'wrapper_function.vba',
    'res2csv.vba',
    'RunPythonInVenv.vba'
]
# open in write mode will automatically clear file contents
with open('./macro.vba', 'w') as outfile:
    for fname in filenames:
        with open('./' + fname, encoding='utf-8') as infile:
            outfile.write(infile.read())

