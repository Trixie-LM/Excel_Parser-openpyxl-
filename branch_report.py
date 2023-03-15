from openpyxl import load_workbook
from numpy import unique
from util import get_cell_value, get_boundary_values


class CommonFunctions:
    def __init__(self):
        self.file_path = 'C:/Users/Trixie_LM/Desktop/1C/Отчет филиала.xlsx'
        self.book = load_workbook(filename=self.file_path, data_only=True)
        self.sheet = self.book.active

    def _get_cell_value(self, column_letter, idx):
        return get_cell_value(column_letter, idx, self.sheet)

class SpreadSheet(CommonFunctions):
    def realization_of_lottery_tickets(self):
        # задаем начальное и конечное значение для поиска диапазона
        start_value_row = 'Наименование лотереи'
        end_value_row = 'ИТОГО:'
        end_value_column = 'Подлежит перечислению Принципалу за отчетный период'

        # координаты начала и конца диапазона
        start_row, start_column, end_row, end_column = None, None, None, None

        # проходим по всем ячейкам листа
        for row in self.sheet.iter_rows():
            for cell in row:
                # если найдено начальное значение, запоминаем его координаты
                if cell.value == start_value_row:
                    start_row, start_column = cell.row+4, cell.column
                # если найдено конечное значение, запоминаем его координаты
                elif cell.value == end_value_row:
                    end_row = cell.row
                # если найдено конечное значение, запоминаем его координаты
                elif cell.value == end_value_column:
                    end_column = cell.column
            if start_row and end_row and end_column:
                break

        # проверяем, что начальное и конечное значение найдены
        if start_row is not None and end_row is not None:
            # определяем диапазон
            range = self.sheet.cell(row=start_row, column=start_column).coordinate + ':' + self.sheet.cell(row=end_row, column=end_column).coordinate
            # выводим найденный диапазон
            return range
        else:
            print('Диапазон не найден')

    def realization_of_lottery_receipts(self):
        # задаем начальное и конечное значение для поиска диапазона
        start_value_row = 'Наименование лотереи'
        end_value_row = 'ИТОГО:'
        end_value_column = 'Подлежит перечислению Принципалу за отчетный период'

        # координаты начала и конца диапазона
        start_row, start_column, end_row, end_column = None, None, None, None

        # проходим по всем ячейкам листа
        for row in self.sheet.iter_rows():
            for cell in row:
                # если найдено начальное значение, запоминаем его координаты
                if cell.value == start_value_row:
                    start_row, start_column = cell.row + 4, cell.column
                # если найдено конечное значение, запоминаем его координаты
                elif cell.value == end_value_row:
                    end_row = cell.row
                # если найдено конечное значение, запоминаем его координаты
                elif cell.value == end_value_column:
                    end_column = cell.column


        # проверяем, что начальное и конечное значение найдены
        if start_row is not None and end_row is not None:
            # определяем диапазон
            range = self.sheet.cell(row=start_row, column=start_column).coordinate + ':' + \
                    self.sheet.cell(row=end_row, column=end_column).coordinate
            # выводим найденный диапазон
            return range
        else:
            print('Диапазон не найден')


class ParsingAndCounting(CommonFunctions):
    def _total_values_in_report(self):
        lottery_tickets_diapason = SpreadSheet().realization_of_lottery_tickets()
        cell_range = self.sheet[lottery_tickets_diapason]


        print(get_boundary_values(lottery_tickets_diapason, 'max_row'))


        lottery_receipts_diapason = SpreadSheet().realization_of_lottery_receipts()
        # for row in cell_range:
        #     for cell in row:
        #         print(cell.value)



ParsingAndCounting()._total_values_in_report()