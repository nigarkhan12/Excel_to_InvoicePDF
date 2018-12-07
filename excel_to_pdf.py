from flask import Flask, render_template
import xlrd
import pdfkit
import os
import shutil, csv
import logging
import time
import pandas as pd
import locale


logging.basicConfig(filename="log_data.log", level=logging.INFO)

project_path = os.getcwd()
wkhtmltopdf_path = os.getcwd() + "\\wkhtmltox\\bin\\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)

locale.setlocale(locale.LC_ALL, 'enn')

excel_path = os.getcwd() + "\\Excel\\Invoice_Generator.xlsx"

img_logo = os.getcwd() + "\\static\\img\\logo.png"
img_seal = os.getcwd() + "\\static\\img\\seal.png"

df_header = pd.read_excel(excel_path).columns.tolist()

header = []

for all_header in df_header:
    header.append(str(all_header))

app = Flask(__name__)


class Data(object):

    def __init__(self, *args, **kwargs):

      valid_keys = ["INV_No", "Ref_No", "Inv_Date", "CIN", "ARN", "Panda_Code",
      "GSTIN", "Unit", "HSN", "Qty","Unit_Price", "IGST", "SGST", "CGST", "Subtotal",
      "Total", "Dealer_Name", "Address1", "Address2", "Address3", "City", 
      "State", "Pincode", "Contact_No", "Description","Description_1", "Rename", 
      "Total_amount_In_Words", "Contact_person", "PO", "PO_Date"]

      final_value = dict(zip(header, args))
         
      self.__dict__.update((k,v) for k, v in final_value.items() if k in valid_keys)
      self.__dict__.update({'img_logo': img_logo , 'img_seal': img_seal})             
      
    def __str__(self):
        return(" Data object: \n"
               "  INV_No = {0}\n"
               "  Ref_No = {1}\n"
               "  Inv_Date = {2}\n"
               "  CIN = {3}\n"
               "  ARN = {4} \n"
               "  Panda_Code = {5} \n"
               "  GSTIN = {6} \n"
               "  Unit = {7} \n"
               "  HSN = {8} \n"
               "  Qty = {9} \n"
               "  Unit_Price = {10} \n"
               "  IGST = {11} \n"
               "  SGST = {12} \n"
               "  CGST = {13} \n"
               "  Subtotal = {14} \n"
               "  Total = {15} \n"
               "  Dealer_Name = {16} \n"
               "  Address1 = {17} \n"
               "  Address2 = {18} \n"
               "  Address3 = {19} \n"
               "  City = {20} \n" 
               "  State = {21} \n"
               "  Pincode = {22} \n"
               "  Contact_No = {23} \n"
               "  Description = {24} \n"
               "  Description_1 = {25} \n"
               "  Rename = {26} \n"
               "  Total_amount_In_Words = {27} \n"
               " Contact_person = {28} \n"
               " PO = {29} \n"
               " PO_Date = {30} \n"
               " img_logo = {31} \n"
               " img_seal = {32} \n"
               .format(self.INV_No , self.Ref_No, self.Inv_Date, self.CIN, self.ARN, self.Panda_Code, 
               self.GSTIN, self.Unit, self.HSN, self.Qty, self.Unit_Price, self.IGST, self.SGST, self.CGST, self.Subtotal, 
               self.Total, self.Dealer_Name, self.Address1, self.Address2, self.Address3, self.City, self.State, 
               self.Pincode, self.Contact_No, self.Description, self.Description_1, self.Rename, self.Total_amount_in_words,
               self.Contact_person, self.PO, self.PO_Date, self.logo, self.seal))



def convertor():
    wb = xlrd.open_workbook(excel_path)
    count = 0
    items = []
    options = {
        'quiet': '',
        'orientation': 'Portrait',
    }
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        rows = []
        for row in (range(1, number_of_rows)):
            values = []
            count += 1
            for col in range(number_of_columns):
                value = (sheet.cell(row, col).value)
                try:
                    if col == 10 or col == 11 or col == 12 or col == 13 or col == 14 or col == 15:
                      value = locale.currency(int(value), grouping=True, symbol
                        =True).replace('?','')
                    else:
                        value = str(int(value))
                except ValueError:
                    pass
                finally:
                    values.append(value)
                if col == 26:
                    file_name = str(value) + ".pdf"

            item = Data(*values).__dict__
            item['id'] = count
            items.append(item)
            with app.app_context():
                style_path = os.getcwd() + "\\static\\assets\\"
                file_dir = os.getcwd() + "\\pdf\\"
                css = os.path.join(style_path, 'style.css')
                html = render_template('index.html', data=item)
                pdfkit.from_string(html, file_name, css=css, options=options, configuration=config)
                source = os.getcwd() + "\\" + file_name
                destination = file_dir + file_name
                shutil.move(source, destination)
    print("Invoice Successfully Generated.")
    return "Success"

if __name__ == '__main__':
    convertor()
