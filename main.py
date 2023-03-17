from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from paid_payments_registry import PreCondition, paid_lottery_names_list
from openpyxl import Workbook
import common_core

print('Создаем файл "Итоговый отчет" и выполняем предусловия...')
file = 'C:/Users/Trixie_LM/Desktop/1C/Итоговый отчет.xlsx'

# Создание файла "Итоговый отчет"
book = Workbook()
sheet = book.active
sheet.title = "Итоговый отчет"

# Создание краткого отчета
PreCondition.copying_table()
PreCondition.delete_rows()

#TODO: добавить в отчет для проверки
# Копирую таблицу "реестр выплаченных выигрышей" из файла в основной
for row in paid_lottery_names_list():
    reversed_row = reversed(row)
    for cell in reversed_row:
    # Перемещаем ячейки из одной таблицы в другую
        sheet.cell(row=cell.row + 2, column=cell.column + 20).value = cell.value


class TestingReports:
    # Редактирование ячеек в файле
    print('Соединяем ячейки и устанавливаем стиль...')
    for cell in common_core.editing_cells():
        double = Side(border_style="double", color="FF000000")
        startCell = cell.split(':')[0]

        sheet[startCell].font = Font(size=10)
        sheet[startCell].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        sheet[startCell].border = Border(top=double, bottom=double, left=double, right=double)
        sheet.merge_cells(cell)

    print('Импортируем данные в таблицу...')
    # Импорт данных в таблицу
    for columnNumber, text in common_core.input_data():
        sheet[columnNumber].value = text


    #TODO: найти решение лучше
    sheet.column_dimensions["F"].auto_size = True
    sheet.column_dimensions["P"].auto_size = True
    sheet.column_dimensions["U"].width = 12
    sheet.column_dimensions["V"].width = 12

    print('Проверяем значения между собой и красим ячейки...')
    # Сверка данных и заливка фона ячейки
    for cellNumber, numberInReport, verifiable in common_core.check_and_painting():
        if str(numberInReport) == str(verifiable):
            sheet[cellNumber].fill = PatternFill('solid', fgColor="00FF00")
        else:
            sheet[cellNumber].fill = PatternFill('solid', fgColor="FF0000")

    book.save(file)


book.save(file)

print("COMPLETE!")
