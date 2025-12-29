from PyPDF2 import PdfReader as PdfFileReader,PdfWriter as PdfFileWriter
import pdfplumber
import sys
import getopt


def usage():
    print(
        """
        delBlankPage.py -i infilePDF -o outfilePDF
        """
    )

def delBlank():
    opts, args = getopt.getopt(sys.argv[1:],"i:o:h",['infile=','outfile=','help='])
    
    for op, value in opts:
        if op == "-i":
            input_file = value
        elif op == "-o":
            output_file = value
        elif op == "-h":
            usage()
            sys.exit()
        else:
            print("Error: invalid parameters")
      
    pdfFileWriter = PdfFileWriter()
    pdfReader = PdfFileReader(input_file)
    pdfRead = pdfplumber.open(input_file)
    
    for index,page in enumerate(pdfRead.pages):
        if page.extract_text() is not None:
            if len(page.extract_text()) > 11:
                pageObj = pdfReader.pages[index]
                pdfFileWriter.add_page(pageObj)
    pdfFileWriter.write(open(output_file, 'wb'))        

if __name__ == "__main__":
    delBlank()
















