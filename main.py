import pandas as  pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for file in filepaths:

    pdf = FPDF(orientation="P",unit="mm",format="A4")
    pdf.add_page()


    filename = Path(file).stem
    invoice_no = filename.split("-")[0]
    date=filename.split("-")[1]

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_no}", align="L", ln=1, border=0)

    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {date}", align="L", ln=1, border=0)
    t_price=0
    dp=pd.read_excel(file,sheet_name="Sheet 1")
    table_columns = list(dp.columns)
    table_columns=[item.replace("_"," ").capitalize() for item in table_columns]
    pdf.set_font(family="Times", size=10,style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt=str(table_columns[0]), border=1)
    pdf.cell(w=70, h=10, txt=str(table_columns[1]), border=1)
    pdf.cell(w=35, h=10, txt=str(table_columns[2]), border=1)
    pdf.cell(w=30, h=10, txt=str(table_columns[3]), border=1)
    pdf.cell(w=30, h=10, txt=str(table_columns[4]), border=1, ln=1)


    for index,row in dp.iterrows():
        pdf.set_font(family="Times",size=10)
        pdf.set_text_color(80,80,80)
        pdf.cell(w=30,h=10, txt=str(row["product_id"]),border=1)
        pdf.cell(w=70,h=10,txt=str(row["product_name"]),border=1)
        pdf.cell(w=35,h=10,txt=str(row["amount_purchased"]),border=1)
        pdf.cell(w=30,h=10,txt=str(row["price_per_unit"]),border=1)
        pdf.cell(w=30,h=10,txt=str(row["total_price"]),border=1,ln=1)
        t_price=t_price+row["total_price"]
    pdf.cell(w=30, h=10, border=1)
    pdf.cell(w=70, h=10, border=1)
    pdf.cell(w=35, h=10, border=1)
    pdf.cell(w=30, h=10, border=1)
    pdf.cell(w=30, h=10, border=1, txt=str(t_price),ln=1)

    pdf.ln(3)
    pdf.set_font(family="Times", size=20, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=0, h=20, txt=f"The total amount due is {t_price} Euros.",ln=1)

    pdf.set_font(family="Times", size=20, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=25, h=20, txt="PythonHow", ln=1)
    pdf.image("pythonhow.png")


    pdf.output(f"PDFs/{filename}.pdf")


