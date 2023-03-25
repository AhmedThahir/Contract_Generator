#!/usr/bin/env python
# coding: utf-8

# In[11]:


import yaml
from fpdf import FPDF
import glob
import concurrent.futures as multi
import os


# In[12]:


font_size = 10
margin = 20
font_family = "Century_Expanded"
line_height = 0.5 * font_size


# In[13]:


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
            
        fonts = glob.glob("fonts/*/*.ttf")
        for font in fonts:
            import_font(font)
    def header(self):
        self.set_y(10)
        image_w = 100
        self.image(
            os.path.join("assets", "header.png"),
            w = image_w,
            x = (self.w - image_w) / 2
        )
        self.set_font(font_family, 'I', font_size)
#         self.cell(0, line_height, "", new_x = "LEFT", new_y="NEXT")
        self.br()

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
            os.path.join("assets", "footer.png"),
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
    
    def save(self, year, agreement, alias_1, alias_2):
        if not os.path.exists("output"):
            os.makedirs("output")

        name = f'{year} {agreement} agreement between {alias_1} and {alias_2}.pdf'
        self.output(
            os.path.join(
                f"output",
                name
            )
        )
        print(f"Saved {name}")


# In[14]:


def gen_contract(contract):
    with open(contract, "r") as f:
        content = yaml.safe_load(f)

    # Create a PDF object
    pdf = PDF('P', 'mm', 'A4')
    pdf.import_fonts()
    pdf.set_margins(top = margin, left = margin, right = margin)
    pdf.set_auto_page_break(auto = True, margin = margin)

    #Add Page
    pdf.add_page()

    pdf.h1("Mutual Benefit Agreement")
    
    pdf.h2("Article 1")
    pdf.p(f"""Between the undersigned
    
**{content["party_1"]}**
represented by {content["representative_1"]}, {content["role_1"]} duly authorized for the purpose hereof, (hereinafter referred to as **{content["alias_1"]}**)

one onehand &

**{content["party_2"]}** whose registered office is located at {content["location_2"]}
represented by {content["representative_2"]}, {content["role_2"]} duly authorized for the purposes hereof, (hereinafter referred to as **{content["alias_2"]}**)

Against this background the following agreement was made:
{" and ".join([content["alias_1"], content["alias_2"]])} have agreed to a Mutually Beneficial Partnership for **{content["agreement"]}**""")
    
    pdf.br()
    
    pdf.h3(f'{content["alias_1"]} to offer {content["alias_2"]}')
    pdf.p(content["party_1_offers_party_2"])
    
    pdf.h3(f'{content["alias_2"]} to offer {content["alias_1"]}')
    pdf.p(content["party_2_offers_party_1"])
    
    pdf.add_page()
    pdf.h2("Article 2")
    pdf.p(f'This agreement shall take effect on commencement of {content["agreement"]} and will be valid till {content["ending"]}. {content["alias_1"]} holds the right to alter the contract and/or extend the validity of the contract, if required.')

    pdf.h3("General Condition")
    pdf.p("In case of cancellation or termination of the agreement, both parties cannot claim compensation from the other party for the product or service delivered before the termination/cancellation of the agreement.")
    
    pdf.br(2)
    pdf.p(f'We agree to the above Terms and Conditions, issued on {content["date_of_issue"]}.')

    pdf.br(2)
    
    for i in range(1, 2+1):
        pdf.p(
f"""**{content[f"representative_{i}"]}**
{content[f"contact_{i}"]}
{content[f"role_{i}"]}, {content[f"alias_{i}"]}
Sign:"""
        )
        pdf.br(3)
   
    pdf.save(content["years"], content["agreement"], content["alias_1"], content["alias_2"])


# In[15]:


contracts = glob.glob("content/*.yml")
for contract in contracts:
    gen_contract(contract)

