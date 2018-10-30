import xlrd
from jinja2 import Environment
from jinja2 import FileSystemLoader


class Data(object):
    ['21AABCG5594P1Z7', '27230', '', 'AMT/AD/2018-19/295', '43383', 'AutoDeal Maintenance Charges', 'From 01-Nov -18 to 30 -Apr-19', '', '', 'Gargson Properties Pvt Ltd.(Capital Ford)', 'A/66, Nayapalli', 'NH 5', '', 'Bhubaneshwar', 'Orissa', '751 003', 'sales@capitalford.net', 'Umesh Ch Panda', '9937231425', '8000', '0', '0', '1440', '9440']

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


wb = xlrd.open_workbook('For Test.xlsx')
for sheet in wb.sheets():
    number_of_rows = sheet.nrows
    number_of_columns = sheet.ncols

    items = []

    rows = []
    for row in range(1, number_of_rows):
        values = []
        for col in range(number_of_columns):
            value = (sheet.cell(row, col).value)
            try:
                value = str(int(value))
            except ValueError:
                pass
            finally:
                values.append(value)
        print(values)
        item = Data(*values)
        items.append(item)
data_query = {
            'GSTIN': items.GSTIN,
            'Panda_Code': items.Panda_Code,
            'PO': items.PO,
            'Inv_No': items.Inv_No,
            'Date': items.Date,
            'Description': items.Description,
            'Description1': items.Description1,
            'Description2': items.Description2,
            'Description3': items.Description3,
            'Dealer_Name': items.Dealer_Name,
            'Address1': items.Address1,
            'Address2': items.Address2,
            'Address3': items.Address3,
            'City': items.City,
            'State': items.State,
            'Pin': items.Pin,
            'Email_Address': items.Email_Address,
            'Contact_person': items.Contact_person,
            'Contact_nos': items.Contact_nos,
            'Base_Amt': items.Base_Amt,
            'CGST': items.CGST,
            'SGST': items.SGST,
            'IGST': items.IGST,
            'Total': items.Total}

for item in items:
    print(item)
    print("Accessing one single value: {0}".format(item.GSTIN))

j2_env = Environment(loader=FileSystemLoader('templates'), trim_blocks=True)

template = j2_env.get_template('new_index.html')

rendered_file = template.render(data_query)
