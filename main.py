import openpyxl
from pyhtml2pdf import converter


# Load the Excel file
workbook = openpyxl.load_workbook('Main.xlsx')

# Get the first sheet in the workbook
sheet = workbook.active

data = []

# iterate over the rows in the sheet
for row in sheet.rows:
    # get the name and website link from the appropriate cells
    fname = row[1].value
    fname = fname.replace(" ","")
    lname = row[2].value
    file_name = fname + lname
    url = row[3]    
    data.append((file_name, url))
    if url.hyperlink:
       url = url.hyperlink.target
       converter.convert(url, file_name+'.pdf')
    print(url , file_name)