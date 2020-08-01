import bs4 as bs
import urllib.request
import xlsxwriter
import re

#Open the webpage and read it in using urllib.request
source = urllib.request.urlopen('https://en.uesp.net/wiki/Online:Shalidor%27s_Library').read()
soup = bs.BeautifulSoup(source,'lxml')

#Grab the div containing all the tables
tables = soup.find('div', {'class': 'mw-content-ltr'})

#Grab all of the table rows
table_rows = tables.find_all('tr')

#Also grab all of the headers and scrip the garbage off of them
table_headers = []
for headers in tables.find_all('span', {'class': 'mw-headline'}):
    table_headers.append(headers.text.strip().replace('[edit]', ''))

#Create an excel workbook and add a sheet
workbook = xlsxwriter.Workbook('checklist.xlsx')
worksheet = workbook.add_worksheet("Shalidor's Library")

#Build a cell formatting object for the headers
header_format = workbook.add_format()
header_format.set_bold()
header_format.set_font_size(16)

#Build a cell formatting object for the cells
cell_format = workbook.add_format(({'text_wrap': True}))
cell_format.set_font_size(14)

#set worksheet column width and format text to wrap
worksheet.set_column(0, 1, 37)
worksheet.set_column(1, 3, 50)
worksheet.set_column(4, 4, 20)

#Build the excel worksheet
table_headers_index = 0
j = 1
table_header_length = len(table_headers)
worksheet.write(0, 0, "Name", header_format)
worksheet.write(0, 1, "Author", header_format)
worksheet.write(0, 2, "Description", header_format)
worksheet.write(0, 3, "Location", header_format)
worksheet.write(0, 4, "Obtained", header_format)
worksheet.freeze_panes(1, 0)
for tr in table_rows:
    td = tr.find_all('td')
    row = [i.text for i in td]
    k = 0
    if not row:
        worksheet.write(j, 0, table_headers[table_headers_index], header_format)
        table_headers_index = table_headers_index + 1
    else:
        for row_data in row:
            worksheet.write(j, k, row_data.strip(), cell_format)
            k = k + 1
    j = j + 1

#Finally, close the workbook (also writes to disk)
workbook.close()