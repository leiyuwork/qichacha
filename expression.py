from bs4 import BeautifulSoup
import os
from openpyxl import load_workbook  # A Python library to read/write Excel xlsx/xlsm files
import difflib

path = r"C:\Users\ray\OneDrive\qichacha\test\\"
files = os.listdir(path)

for file in files:
    result = []
    if 'null' in file:
        result.append(file.replace('.html', '').replace('null', ''))
        result.append('null')
        result.append('null')
        result.append('null')
    else:
        try:
            bs_obj = BeautifulSoup(open(path + file, encoding='UTF-8'), 'html.parser')

            title = bs_obj.select('title')[0].get_text().replace(' - 企查查', '')
            og_type = (bs_obj.select('td')[25].get_text().strip())
            area = bs_obj.select('td')[35].get_text().strip()

            result.append(file.replace('.html', ''))
            result.append(og_type)
            result.append(area)
            result.append(title)
            result.append(difflib.SequenceMatcher(file.replace('.html', ''), title).ratio())

        except:
            bs_obj = BeautifulSoup(open(path + file, encoding='gbk'), 'html.parser')

            title = bs_obj.select('title')[0].get_text().replace(' - 企查查', '')
            og_type = (bs_obj.select('td')[25].get_text().strip())
            area = bs_obj.select('td')[35].get_text().strip()

            result.append(file.replace('.html', ''))
            result.append(og_type)
            result.append(area)
            result.append(title)
            result.append(difflib.SequenceMatcher(file.replace('.html', ''), title).ratio())
    print(result)

    wb = load_workbook(r'C:\Users\ray\OneDrive\qichacha\result.xlsx')  # open excel
    sheet = wb.active  # get the currently active sheet
    sheet.append(result)  # append matching result to excel sheet
    wb.save(r'C:\Users\ray\OneDrive\qichacha\result.xlsx')  # save excel

"""
    lis = bs_obj.find_all('div', {"class": "searchCompanyName"})
    for a in lis:
        result = []
        result.append(a.text)
        result.append(a.find(name='a').get('href').replace('/company.php?m_id=', ''))
        print(result)

"""


