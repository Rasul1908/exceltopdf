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
    invoice_nr,invoice_date =filename.split('-')

    ### invoice no
    pdf.set_font(family='Times',size= 20,style='B')
    pdf.cell(w=50,h=8,txt=f'Invoice Nr {invoice_nr} Date :{invoice_date}',align='L',ln=1)

    ### Adding modified column names
    columns=list(df.columns)
    columns=[item.replace('_', ' ').capitalize() for item in columns]
    pdf.set_font(family='Times',size= 10,style='I')
    pdf.cell(w=35,h=12,txt=columns[0],border=1)
    pdf.cell(w=35,h=12,txt=columns[1],border=1)
    pdf.cell(w=35,h=12,txt=columns[2],border=1)
    pdf.cell(w=35,h=12,txt=columns[3],border=1)
    pdf.cell(w=35,h=12,txt=columns[4],border=1,ln=1)

    for index,row in df.iterrows():
        pdf.set_font(family='Times', size=10, style='I')
        pdf.cell(w=35,h=12,txt=str(row['product_id']),border=1)
        pdf.cell(w=35,h=12,txt=str(row['product_name']),border=1)
        pdf.cell(w=35,h=12,txt=str(row['amount_purchased']),border=1)
        pdf.cell(w=35,h=12,txt=str(row['price_per_unit']),border=1)
        pdf.cell(w=35,h=12,txt=str(row['total_price']),border=1,ln=1)

    pdf.set_font(family='Times', size=10, style='I')
    pdf.cell(w=35, h=12, txt= '', border=0)
    pdf.cell(w=35, h=12, txt= '', border=0)
    pdf.cell(w=35, h=12, txt= '', border=0)
    pdf.set_font(family='Times', size=10, style='B')
    pdf.cell(w=35, h=12, txt= 'Grand Total', border=1)
    pdf.cell(w=35,h=12,txt=str(df['total_price'].sum()),border=1,ln=1)

    pdf.set_font(family='Times', size=14, style='B')
    pdf.cell(w=35,h=12,txt=str(f"Total amount of the invoice is {df['total_price'].sum()} EUR"),ln=1)
    pdf.set_font(family='Times', size=14, style='I')
    pdf.cell(w=25,h=8,txt="",ln=1)

    pdf.cell(w=25,h=8,txt="Rasul_Python2023",ln=1)
    pdf.image('pythonhow.png',w=10)



    pdf.output(f'PDFS/{filename}.pdf')


