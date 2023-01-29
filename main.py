import pandas
import glob
from fpdf import FPDF
import pathlib

df=pandas.read_excel('EXCELFILES/10001-2023.1.18.xlsx')

filepaths=glob.glob('EXCELFILES/*.xlsx')

for filepath in filepaths:
    df=pandas.read_excel(filepath,sheet_name='Sheet 1')

    pdf = FPDF(orientation='P', format='a4', unit='mm')
    pdf.add_page()

    ### intel string defines filepath/filename
    filename = pathlib.Path(filepath).stem
    invoice_nr=filename.split('-')[0]
    invoice_date=filename.split('-')[1]

    ### invoice no
    pdf.set_font(family='Times',size= 16,style='B')
    pdf.cell(w=50,h=8,txt=f'Invoice Nr {invoice_nr} Date :{invoice_date}')




    pdf.output(f'PDFS/{filename}.pdf')


