from flask import Flask, render_template
import xlrd
from xhtml2pdf import pisa
from io import StringIO, BytesIO
import pdfkit
import os
import shutil

app = Flask(__name__)


class Data(object):

    def __init__(self, GSTIN, Panda_Code, PO, Inv_No, Date, Description, Description1, Description2, Description3, Dealer_Name, Address1, Address2, Address3, City, State, Pin, Email_Address, Contact_person, Contact_nos, Base_Amt, CGST, SGST, IGST, Total):
        self.id = id
        self.GSTIN = GSTIN
        self.Panda_Code = Panda_Code
        self.PO = PO
        self.Inv_No = Inv_No
        self.Date = Date
        self.Description = Description
        self.Description1 = Description1
        self.Description2 = Description2
        self.Description3 = Description3
        self.Dealer_Name = Dealer_Name
        self.Address1 = Address1
        self.Address2 = Address2
        self.Address3 = Address3
        self.City = City
        self.State = State
        self.Pin = Pin
        self.Email_Address = Email_Address
        self.Contact_person = Contact_person
        self.Contact_nos = Contact_nos
        self.Base_Amt = Base_Amt
        self.CGST = CGST
        self.SGST = SGST
        self.IGST = IGST
        self.Total = Total

    def __str__(self):
        return(" Data object: \n"
               "  Data_id = {0}\n"
               "  GSTIN = {1}\n"
               "  Panda_Code = {2}\n"
               "  PO = {3}\n"
               "  Inv_No = {4} \n"
               "  Date = {5} \n"
               "  Description = {6} \n"
               "  Description1 = {7} \n"
               "  Description2 = {8} \n"
               "  Description3 = {9} \n"
               "  Dealer_Name = {10} \n"
               "  Address1 = {11} \n"
               "  Address2 = {12} \n"
               "  Address3 = {13} \n"
               "  City = {14} \n"
               "  State = {15} \n"
               "  Pin = {16} \n"
               "  Email_Address = {17} \n"
               "  Person = {18} \n"
               "  Contact_nos = {19} \n"
               "  Base_Amt = {20} \n"
               "  CGST = {21} \n"
               "  SGST = {22} \n"
               "  IGST = {23} \n"
               "  Total = {24} \n"
               .format
               (self.id, self.GSTIN, self.Panda_Code,
               self.PO, self.Inv_No, self.Date, self.Description, self.Description1, self.Description2, self.Description3,
               self.Dealer_Name, self.Address1, self.Address2, self.Address3, self.City, self.State, self.Pin, self.Email_Address,
               self.Contact_person, self.Contact_nos, self.Base_Amt, self.CGST, self.SGST, self.IGST, self.Total))


def convertor(id):
    wb = xlrd.open_workbook('For Test.xlsx')
    count = 0
    items = []
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
        rows = []
        for row in range(1, number_of_rows):
            values = []
            count += 1
            if count == int(id):
                for col in range(number_of_columns):
                    value = (sheet.cell(row, col).value)
                    try:
                        value = str(int(value))
                    except ValueError:
                        pass
                    finally:
                        values.append(value)
                item = Data(*values).__dict__
                item['id'] = count
                items.append(item)
    return items


@app.route("/<id>")
def index(id):
    data = convertor(int(id))[0]
    file_name = "invoice-" + id + ".pdf"
    options = {
        'page-size': 'A4',
        'margin-top': '0.75in',
        'margin-right': '0.75in',
        'margin-bottom': '0.75in',
        'margin-left': '0.75in',
    }
    css = os.path.join('/home/nigar/Desktop/Excel_to_pdf/static/assets/', 'style.css')
    html = render_template('wolters/index.html', data=data)
    pdfkit.from_string(html, file_name, css=css)
    source = '/home/nigar/Desktop/Excel_to_pdf/' + file_name
    destination = '/home/nigar/Desktop/Excel_to_pdf/pdf/' + file_name
    shutil.move(source, destination)
    return render_template('wolters/index.html', data=data)


if __name__ == '__main__':
    app.run(debug=True)
