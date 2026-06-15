from pypdf import PdfWriter
import glob

# List all .pdf files in the pdf directory
pdf_files = glob.glob(".\pdf\*.pdf")

# read list from no_paper.txt
with open("no_paper.txt", "r") as file:
    no_paper_list = [line.strip() for line in file if line.strip()]

filtered_list = [
    item for item in pdf_files 
    if not any(sub in item for sub in no_paper_list)
]

# define merger as pdf writer
merger = PdfWriter()

for pdf in filtered_list:
    merger.append(pdf)

# Write the final result to a new file
merger.write("all.pdf")
merger.close()
