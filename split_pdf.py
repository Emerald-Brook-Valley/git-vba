from pypdf import PdfReader, PdfWriter
import argparse

def split_pdf(input_path, uploadList):
    reader = PdfReader(input_path)
    # Iterate through every page in the PDF
    oldName = ""
    repeat = 0
    for i, page in enumerate(reader.pages):
        writer = PdfWriter()
        writer.add_page(page)
        fileName = uploadList[i]
        fileName = fileName.replace("\n", "")
        if fileName != oldName:
            with open(f"pdf/{fileName}.pdf", "wb") as f:
                writer.write(f)
            repeat = 0
        else:
            repeat = repeat + 1
            with open(f"pdf/{fileName}-{repeat}.pdf", "wb") as f:
                writer.write(f)
        oldName = fileName

def main():
    # listFile.csv 存取pdf對應戶名
    with open('listFile.csv', 'r') as f:
        uploadList = [line.strip() for line in f]
    
    parser = argparse.ArgumentParser(description="split pdf script")
    parser.add_argument("wholePdf", help="File of whole pdf") # Positional
    parser.add_argument("-v", "--verbose", action="store_true") # Flag

    args = parser.parse_args()
    split_pdf(args.wholePdf, uploadList)

if __name__ == '__main__':
    main()
    