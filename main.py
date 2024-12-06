import pandas as  pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for file in filepaths:
    dp=pd.read_excel(file,sheet_name="Sheet 1")
    pdf = FPDF(orientation="P",unit="mm",format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=16)
    filename = Path(file).stem
    invoice_no = filename.split("-")[0]
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_no}", align="L", ln=1, border=0)
    pdf.output(f"PDFs/{filename}.pdf")


