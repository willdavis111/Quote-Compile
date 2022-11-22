import PyPDF2
import re
import glob
from openpyxl import workbook, load_workbook
xl_path = r'D:\Quote Compilation\quote_compile\Excell\quote_compile.xlsx'
wb = load_workbook(xl_path)
pipe_sheet = wb['PIPE']
fitting_sheet = wb['FITTINGS']


lines = []
pipe_item = []
pipe_item_price = []
fitting_item = []
fitting_item_price = []

#Reads all pdfs in the folder bellow
def compile_folder():
    quote_folder = r'D:\Quote Compilation\quote_compile\core_quotes'
    for file in glob.glob(quote_folder + r"\*"):
        extract_pdf_text(file)

#Makes a string of all text in all files
def extract_pdf_text(file):
    with open(file, 'rb') as pdfFile:
        reader = PyPDF2.PdfFileReader(pdfFile)
        number_of_pages = len(reader.pages)
        count = -1
        for pages in range(number_of_pages):
            count += 1
            page = reader.getPage(count)
            text = page.extractText()
            lines1 = text.splitlines()
            item_price_extract(" EA ", lines1)
            item_price_extract(" FT ", lines1)


#extracts desired infor(product, price) creates two lists, one of products one of prices
def item_price_extract(unit, line_data):
    global fitting_item_price, pipe_item_price, fitting_item, pipe_item, item_actual
    fittings_lines = []
    for line in line_data:
        if re.search(unit, line):
            fittings_lines.append(line)
    for price_full in fittings_lines:
        price_and_total = price_full.split(unit)[1]
        item_and_serial = price_full.split(unit)[0]
        price = price_and_total.split(r' ')[0]
        if price[0].isdigit():
            quantity = float(price_and_total.split(r' ')[1].replace(',', '')) / float(price.replace(',', ''))
            item_and_quantity = item_and_serial.split(r' ', 1)[1]
            test_quantity_value = item_and_quantity.split(r' ', 1)[0]
            if test_quantity_value == quantity:
                item_actual = item_and_quantity.split(r' ', 1)[1]
            elif test_quantity_value != quantity:
                item_actual = item_and_quantity[(len(str(int(quantity)))):]
            if item_actual[0] == " ":
                item_actual = item_actual.lstrip()
            if unit == ' FT ':
                pipe_item.append(item_actual)
                pipe_item_price.append(price)
            elif unit == ' EA ':
                fitting_item.append(item_actual)
                fitting_item_price.append(price)

#creates a dictionary of unique products and all corresponding prices
def product_price_dict(list2, list3):
    global dict1
    dict1 = {}
    for item in set(list2):
        indices = [i for i, x in enumerate(list2) if x == item]
        listx = []
        for index in indices:
            listx.append(list3[index])
        dict1[item] = listx


#looks for keywords in products to assign material
def assign_material(item):
    global part_material
    pvc = ['PVC', 'SDR25', 'SDR35', 'MOLDED', 'CLEANOUT']
    clay = ['CLAY', 'LOGAN']
    water = ['NIP', 'TUB', 'COPPER', 'MJ', 'DI', 'FLG', 'IMP']
    storm = ['HDPE', 'N12', 'HP']
    if any(substring in item for substring in pvc):
        part_material = "PVC"
    # elif any(substring in item for substring in dip):
    #     part_material = "DIP"
    elif any(substring in item for substring in storm):
        part_material = "STM"
    elif any(substring in item for substring in water):
        part_material = "water"
    elif any(substring in item for substring in clay):
        part_material = "CLAY"
    else:
        part_material = "MISC."


#transfers the created dictionaries to an excel document
def xl_transfer(dict_price, sheet):
    count3 = 1
    for x, prices_list in dict_price.items():
        count3 += 1
        count_cols = 7
        assign_material(x)
        sheet['G' + str(count3)] = x
        sheet['F' + str(count3)] = part_material
        for s in prices_list:
            count_cols += 1
            s = s.replace(',', '')
            sheet.cell(row=count3, column=count_cols).value = float(s)
    wb.save(xl_path)

run = input("Update Core Quotes(y,n):  ")
if run == "y":
    compile_folder()
    product_price_dict(pipe_item, pipe_item_price)
    xl_transfer(dict1, pipe_sheet)
    product_price_dict(fitting_item, fitting_item_price)
    xl_transfer(dict1, fitting_sheet)

