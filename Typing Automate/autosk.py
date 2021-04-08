import glob

from PyPDF2 import PdfFileWriter, PdfFileReader
from fpdf import FPDF
from datetime import datetime
import os

studentroll=None

def merger(output_path, input_paths):
    pdf_writer = PdfFileWriter()
 
    for path in input_paths:
        pdf_reader = PdfFileReader(path)
        for page in range(pdf_reader.getNumPages()):
            pdf_writer.addPage(pdf_reader.getPage(page))
 
    with open(output_path, 'wb') as fh:
        pdf_writer.write(fh)

        

def createcerti():
    from fpdf import FPDF
 
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    
    pdf.add_page()

    pdf.set_font("Arial", size=12)
    pg=1
    #dt=input("Enter Date in format 1/28/2020")
    dt="1/28/2020"

    now = datetime.now()
    dat = now.strftime("%m/%d/%Y")
    tim = now.strftime("%H:%M:%S")
    print(dat,tim)
    s="Typing Test completed | "+dt+" | Typing Master                            Page  "+ str(pg)
    pdf.cell(200, 10, txt=s, ln=1, align="L")
    pdf.set_line_width(1)
    pdf.line(10, 20, 180, 20)

    pdf.set_font('Arial', 'B', 8)
    
    pdf.cell(200, 7, txt="TYPING TEST - PASSED", ln=1, align="L")

    s="User:            1234567"
    pdf.cell(0, 5, txt=s, ln=1, align="L")
    s="Test Name:       DNA research"
    pdf.cell(0, 5, txt=s, ln=1, align="L")
    s="Date :           1/28/2020"
    pdf.cell(0, 5, txt=s, ln=1, align="L")

    pdf.cell(0, 7, txt="", ln=1, align="L")
    
    
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(200, 7, txt="TEST RESULTS", ln=1, align="L")

    s="Duration:        1:23 min of total time 10 :00 min"
    pdf.cell(0, 5, txt=s, ln=1, align="L")
    s="Gross Speed:     74 wpm           Gross strokes:   74"
    pdf.cell(0, 5, txt=s, ln=1, align="L")
    s="Accuracy:        67%              Error hits:      165(33 errors*5)"
    pdf.cell(0, 5, txt=s, ln=1, align="L")
    s="Net Speed:       60 wpm           Netstrokes       91"
    pdf.cell(0, 5, txt=s, ln=1, align="L")

    pdf.cell(0, 7, txt="", ln=1, align="L")
    pdf.cell(200, 7, txt="TEST TEXT", ln=1, align="L")

    file=open("textpara\para1.exm")
    s=file.read(800)
    file.close()
    print(s)
    s=str(s)
    pdf.set_font('Arial', '', 8)
    pdf.set_margins(left=0.0, top=0.0, right=50.0)
    #pdf.cell(200, 5, txt=s, ln=1, align="L")
    pdf.multi_cell(w=0, h=5, txt=s, border=0)
    

    
    pdf.output("simple_demo.pdf")
    print("Pdf Creation Complete ")

def readexcel():
    data="ExcelRollno.xls"
    # Reading an excel file using Python 
    import xlrd 
    loc = (data) 
    wb = xlrd.open_workbook(loc) 
    sheet = wb.sheet_by_index(0)
    
    for i in range(1,sheet.nrows):
        print(int(sheet.cell_value(i, 0)))

        p=os.getcwd()+"\set_1\*.pdf"
        print(p)
        paths = glob.glob(p)
        paths.sort()
        fnm=str(int(sheet.cell_value(i, 0)))+".pdf"
        merger(fnm, paths)
        print("Done")





createcerti()
readexcel()








