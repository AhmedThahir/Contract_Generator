# -*- coding: utf-8 -*-
"""Contract_Generator.ipynb

Automatically generated by Colaboratory.

Original file is located at
    https://colab.research.google.com/drive/1PnH5lgo01YUILpSrfWA_iXByjNyF9cvs
"""

!pip install fpdf2
from fpdf import FPDF
import os
from pathlib import Path
import pandas as pd
import glob

from google.colab import drive
drive.mount('/content/drive/')
folder = "/content/drive/MyDrive/BMS/Sponsorships/Contracts/Contract_Generator"

font_size = 10
margin_factor = 1/8 # fraction of page width for margin on each side
font_family = "Century_Expanded"
line_height = 1/2 * font_size

sheet = pd.read_excel(
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vStG1HTCIkLF8olKOgLQ-4Gz6fSma8hBgMwr_IHNUt-YUJtei-uLZ0x4ZNZkge9F7BHOgJeF1wJnWaY/pub?output=xlsx",
    sheet_name = "Details"
)
sheet

class PDF(FPDF):
    def import_fonts(self):
        def import_font(fname):
            family = os.path.basename(os.path.dirname(fname))
            style = ""
            if "bold" in fname.lower():
                style+="B"
            if "italic" in fname.lower():
                style+="I"
            
            self.add_font(family=family, fname = fname, style=style)
            
        fonts = glob.glob(
            os.path.join(
              folder,
              "fonts/*/*.ttf"
            )
        )
        for font in fonts:
          import_font(font)
    
    def margin(self, margin_factor):
      margin = margin_factor * self.w
      self.set_margins(top = margin, left = margin, right = margin)
      self.set_auto_page_break(auto = True, margin = margin)

    def header(self):
        self.set_y(5)
        image_w = 100
        self.image(
            os.path.join(folder, "assets/header.png"),
            w = image_w,
            x = (self.w - image_w) / 2
        )
        self.set_font(font_family, 'I', font_size)
#         self.cell(0, line_height, "", new_x = "LEFT", new_y="NEXT")

    def footer(self):
        self.set_y(-30)
        
        # set font
        self.set_font(font_family, '', font_size)
        # Page number
        self.alias_nb_pages()
        self.cell(0, line_height, f'Page {self.page_no()}/{{nb}}', align='C')      
        
        self.set_y(-20)
        image_w = 140
        self.image(
            os.path.join(folder, "assets/footer.png"),
            w = image_w,
            x = (self.w - image_w) / 2
        )
        
    def h1(self, text):
        factor = 1.2
        self.set_font(font_family, 'BU', factor*font_size)
        self.cell(0, 1.5*factor*line_height, text, align='C', new_x = "LEFT", new_y="NEXT")
    def h2(self, text):
        factor = 1.2
        self.set_font(font_family, 'B', factor*font_size)
        self.cell(0, 1.5*factor*line_height, text, align='C', new_x = "LEFT", new_y="NEXT")
    def h3(self, text):
        factor = 1
        self.set_font(font_family, 'BU', factor*font_size)
        self.cell(0, 1.5*factor*line_height, text, new_x = "LEFT", new_y="NEXT")
    def p(self, text):
        self.set_font(font_family, '', font_size)
        self.multi_cell(0, line_height, text, markdown = True, new_x = "LEFT", new_y="NEXT")
    
    def br(self, count=1):
        self.cell(0, count*line_height, new_x = "LEFT", new_y="NEXT")
    
    def save(self, season, agreement, alias_1, alias_2):
        path = Path(folder)
        parent = path.parent.absolute()
        
        output_folder = os.path.join(parent, f"Generated_Contracts/{season}")
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        name = f'{season} {agreement} agreement between {alias_1} and {alias_2}.pdf'
        self.output(
            os.path.join(
                output_folder,
                name
            )
        )
        print(f"Saved {name}")

        # make it easier for uploading signed contracts
        output_folder = os.path.join(parent, f"Signed_Contracts/{season}")
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

def gen_contract(contract):
    if "agreement" not in contract:
      contract["agreement"] = "BITS Dubai Motorsports Project"
    if "party_1" not in contract:
      contract["party_1"] = "Birla Institute of Technology and Science, Pilani - Dubai Campus"
    if "alias_1" not in contract:
      contract["alias_1"] = "BPDC"
      

    pdf = PDF('P', 'mm', 'A4')
    pdf.import_fonts()
    pdf.margin(margin_factor)

    #Add Page
    pdf.add_page()

    pdf.h1("Mutual Benefit Agreement")
    
    pdf.h2("Article 1")
    pdf.p(f"""Between the undersigned
    
**{contract["party_1"]}** represented by {contract["representative_1"]}, {contract["role_1"]} duly authorized for the purpose hereof, (hereinafter referred to as **{contract["alias_1"]}**)

one onehand &

**{contract["party_2"]}** whose registered office is located at {contract["location_2"]} represented by {contract["representative_2"]}, {contract["role_2"]} duly authorized for the purposes hereof (hereinafter referred to as **{contract["alias_2"]}**).

Against this background the following agreement was made: {contract["alias_1"]} and {contract["alias_2"]} have agreed to a Mutually Beneficial Partnership for **{contract["agreement"]}**.""")
    
    pdf.br()
    
    pdf.h3(f'{contract["alias_1"]} to offer {contract["alias_2"]}')
    pdf.p(contract["party_1_offers_party_2"])
    
    pdf.h3(f'{contract["alias_2"]} to offer {contract["alias_1"]}')
    pdf.p(contract["party_2_offers_party_1"])

    pdf.add_page()
    
    pdf.h2("Article 2")
    pdf.p(f'This agreement shall take effect on commencement of {contract["agreement"]} and will be valid till {contract["ending"]}. {contract["alias_1"]} holds the right to alter the contract and/or extend the validity of the contract, if required.')

    pdf.h3("General Condition")
    pdf.p("In case of cancellation or termination of the agreement, both parties cannot claim compensation from the other party for the product or service delivered before the termination/cancellation of the agreement.")
    
    pdf.br(2)
    pdf.p(f'We agree to the above Terms and Conditions, issued on {contract["date_of_issue"].strftime("%b %d, %Y")}.')
    pdf.br(2)
    
    for i in range(1, 2+1):
        pdf.p(
f"""**{contract[f"representative_{i}"]}**
{contract[f"contact_{i}"]}
{contract[f"role_{i}"]}, {contract[f"alias_{i}"]}
Sign:"""
        )
        pdf.br(3)
   
    pdf.save(contract["season"], contract["agreement"], contract["alias_1"], contract["alias_2"])

def check(row):
  missing = row.isnull().values.any()
  if missing:
    print("Missing value for", row, sep="\n\n")
    raise Exception

def clean_df(data):
  df = data.copy()
  for column in df.columns:
    if "date" in column:
      df[column] = pd.to_datetime(sheet[column])
  df = df.apply(lambda col: col.str.strip() if col.dtype == "object" else col)
  return df

sheet.apply(check, axis=1)
sheet = sheet.pipe(clean_df)

sheet.apply(gen_contract, axis = 1)

print("""
Done :)
View the output in: https://drive.google.com/drive/folders/15qEgGZ4YKmr9jUXUZXoM7SXVkg2Wzw6T
""")