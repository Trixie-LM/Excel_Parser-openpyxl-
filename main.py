from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl import Workbook
import common_core

file = 'C:/Users/Trixie_LM/Desktop/1C/Итоговый отчет.xlsx'

# Создание файла "Итоговый отчет"
excel = Workbook()
sheet = excel.active
sheet.title = "Итоговый отчет"


class TestingReports:
    # Редактирование ячеек в файле
    for cell in common_core.editing_cells():
        double = Side(border_style="double", color="FF000000")
        startCell = cell.split(':')[0]

        sheet[startCell].font = Font(size=10)
        sheet[startCell].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        sheet[startCell].border = Border(top=double, bottom=double, left=double, right=double)
        sheet.merge_cells(cell)

    # Импорт данных в таблицу
    for columnNumber, text in common_core.input_data():
        sheet[columnNumber].value = text

    sheet.column_dimensions["F"].auto_size = True
    sheet.column_dimensions["P"].auto_size = True

    # Сверка данных и заливка фона ячейки
    for cellNumber, numberInReport, verifiable in common_core.check_and_painting():
        if str(numberInReport) == str(verifiable):
            sheet[cellNumber].fill = PatternFill('solid', fgColor="00FF00")
        else:
            sheet[cellNumber].fill = PatternFill('solid', fgColor="FF0000")

    excel.save(file)

print("COMPLETE!")
