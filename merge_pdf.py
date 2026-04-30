from pypdf import PdfWriter
import glob

# List all .pdf files in the pdf directory
pdf_files = glob.glob(".\pdf\*.pdf")
merger = PdfWriter()

for pdf in pdf_files:
    merger.append(pdf)

# Write the final result to a new file
merger.write("all.pdf")
merger.close()
