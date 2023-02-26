from openpyxl import load_workbook

document = 'C:/Users/Trixie_LM/Desktop/ExcelParser/Отчет для сверки.xlsm'

book = load_workbook(filename=document, data_only=True)
sheet = book.active

# Продажа
def number_of_circulation_sales():
    return sheet['E' + str(sheet.max_row)].value

def amount_of_circulation_sales():
    return sheet['F' + str(sheet.max_row)].value

def number_of_digital_sales():
    return sheet['G' + str(sheet.max_row)].value

def amount_of_digital_sales():
    return sheet['H' + str(sheet.max_row)].value

def number_of_instant_sales():
    return sheet['I' + str(sheet.max_row)].value

def amount_of_instant_sales():
    return sheet['J' + str(sheet.max_row)].value

def total_number_of_sales():
    return sheet['K' + str(sheet.max_row)].value

def total_amount_of_sales():
    return sheet['L' + str(sheet.max_row)].value

# Выплата
def number_of_circulation_payments():
    return sheet['M' + str(sheet.max_row)].value

def amount_of_circulation_payments():
    return sheet['N' + str(sheet.max_row)].value

def number_of_digital_payments():
    return sheet['O' + str(sheet.max_row)].value

def amount_of_digital_payments():
    return sheet['P' + str(sheet.max_row)].value

def number_of_instant_payments():
    return sheet['Q' + str(sheet.max_row)].value

def amount_of_instant_payments():
    return sheet['R' + str(sheet.max_row)].value

def total_number_of_payments():
    return sheet['S' + str(sheet.max_row)].value

def total_amount_of_payments():
    return sheet['T' + str(sheet.max_row)].value
