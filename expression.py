import bs4
import os
import re
import lxml
from openpyxl import load_workbook  # A Python library to read/write Excel xlsx/xlsm files
import difflib

path = r"C:\Users\ray94\OneDrive\qichacha\4001-6000\\"
files = os.listdir(path)

for file in files:
    result = []
    if 'null' in file:
        result.append(file.replace('.html', '').replace('null', ''))
        result.append('null')
        result.append('null')
        result.append('null')
        result.append('0')
    else:
        try:
            bs_obj = bs4.BeautifulSoup(open(path + file, encoding='UTF-8'), 'lxml')
            try:

                title = bs_obj.select('title')[0].get_text().replace(' - 企查查', '')

                bs_obj_str = str(bs_obj)
                #print(bs_obj_str)

                pattern_og = re.compile(r'企业类型.{1,50}\s.{1,100}')
                find_og = re.findall(pattern_og, bs_obj_str)
                find_og = [c.replace('企业类型</td> <td class="" width="20%">', "").strip() for c in find_og]
                og_type = find_og[0]

                pattern_ar = re.compile(r'所属地区.{1,50}\s.{1,100}')
                find_ar = re.findall(pattern_ar, bs_obj_str)
                find_ar = [c.replace('所属地区</td> <td class="">', "").strip() for c in find_ar]
                area = find_ar[0]

                result.append(file.replace('.html', ''))
                result.append(og_type)
                result.append(area)
                result.append(title)
                result.append(difflib.SequenceMatcher(None, file.replace('.html', ''), title).ratio())
            except:
                result.append(file.replace('.html', '').replace('null', ''))
                result.append('null')
                result.append('null')
                result.append('null')
                result.append('0')
        except:
            bs_obj = bs4.BeautifulSoup(open(path + file, encoding='gbk'), 'lxml')
            try:

                title = bs_obj.select('title')[0].get_text().replace(' - 企查查', '')

                bs_obj_str = str(bs_obj)
                # print(bs_obj_str)

                pattern_og = re.compile(r'企业类型.{1,50}\s.{1,100}')
                find_og = re.findall(pattern_og, bs_obj_str)
                find_og = [c.replace('企业类型</td> <td class="" width="20%">', "").strip() for c in find_og]
                og_type = find_og[0]

                pattern_ar = re.compile(r'所属地区.{1,50}\s.{1,100}')
                find_ar = re.findall(pattern_ar, bs_obj_str)
                find_ar = [c.replace('所属地区</td> <td class="">', "").strip() for c in find_ar]
                area = find_ar[0]

                result.append(file.replace('.html', ''))
                result.append(og_type)
                result.append(area)
                result.append(title)
                result.append(difflib.SequenceMatcher(None, file.replace('.html', ''), title).ratio())
            except:
                result.append(file.replace('.html', '').replace('null', ''))
                result.append('null')
                result.append('null')
                result.append('null')
                result.append('0')
    print(result)

    wb = load_workbook(r'C:\Users\ray94\Desktop\4001-6000shanshan.xlsx')  # open excel
    sheet = wb.active  # get the currently active sheet
    sheet.append(result)  # append matching result to excel sheet
    wb.save(r'C:\Users\ray94\Desktop\4001-6000shanshan.xlsx')  # save excel



