from openpyxl import load_workbook
from util import _get_cell_value, get_boundary_values


class CommonFunctions:
    def __init__(self):
        self.file_path = 'C:/Users/Trixie_LM/Desktop/1C/Отчет агента для бестиражных лотерей.xlsx'
        self.book = load_workbook(filename=self.file_path, data_only=True)
        self.sheet = self.book.active

    def _get_cell_value(self, column_letter, idx):
        return _get_cell_value(column_letter, idx, self.sheet)


class WorkSheet(CommonFunctions):
    def realization_of_lottery_tickets(self):
        # задаем начальное и конечное значение для поиска диапазона
        start_value_row = 'Наименование лотереи'
        end_value_row = 'ИТОГО:'
        end_value_column = 'Остаток на конец  отчетного периода'

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
            if start_row and end_row and end_column:
                break

        # проверяем, что начальное и конечное значение найдены
        if start_row is not None and end_row is not None:
            # определяем диапазон
            range = self.sheet.cell(row=start_row, column=start_column).coordinate + ':' + self.sheet.cell(row=end_row,
                                                                                                           column=end_column).coordinate
            # выводим найденный диапазон
            return range
        else:
            print('Диапазон не найден')


# Объекты класса WorkSheet
tickets_table_range = WorkSheet().realization_of_lottery_tickets()


class ReportAgentNoncirculatedData(CommonFunctions):
    # Беру данные из "ИТОГО" в таблице "Реализация лотерейных билетов"
    def total_values_lottery_tickets(self, column):
        result_row = get_boundary_values(tickets_table_range, 'max_row')
        # Продажи
        sold_number = self._get_cell_value('H', result_row)
        sold_amount = self._get_cell_value('I', result_row)
        # Выплаты
        paid_number = self._get_cell_value('M', result_row)
        paid_amount = self._get_cell_value('N', result_row)
        # Вознаграждения
        reward = self._get_cell_value('P', result_row)
        # Перечисление принципалу
        transfer = self._get_cell_value('Q', result_row)

        if column == 'sold_number':
            return sold_number
        elif column == 'sold_amount':
            return sold_amount
        elif column == 'paid_number':
            return paid_number
        elif column == 'paid_amount':
            return paid_amount
        elif column == 'reward':
            return reward
        elif column == 'transfer':
            return transfer

    # Итоговые суммы ниже таблицы
    def reward_under_table(self):
        result_row = int(get_boundary_values(tickets_table_range, 'max_row')) + 8
        rubles = self._get_cell_value('G', result_row)
        pennies = self._get_cell_value('J', result_row)/100
        reward = rubles + pennies
        return reward

    def transfer_under_table_to_principal(self):
        result_row = int(get_boundary_values(tickets_table_range, 'max_row')) + 11
        rubles = self._get_cell_value('G', result_row)
        pennies = self._get_cell_value('J', result_row)/100
        reward = rubles + pennies
        return reward

    def transfer_under_table_to_agent(self):
        result_row = int(get_boundary_values(tickets_table_range, 'max_row')) + 14
        rubles = self._get_cell_value('G', result_row)
        pennies = self._get_cell_value('J', result_row)/100
        reward = rubles + pennies
        return reward


class AgentNoncirculatedAsserts(CommonFunctions):
    # Функция для подсчета итоговых данных
    # в таблице "Реализация лотерейных билетов"
    def counting_values_in_column(self, column, data_type='int'):
        total = 0
        min_row = int(get_boundary_values(tickets_table_range, 'min_row'))
        max_row = int(get_boundary_values(tickets_table_range, 'max_row'))

        for row in range(min_row, max_row):
            cell = self.sheet.cell(row=row, column=column).value
            if data_type == 'float':
                total += float(cell)
            else:
                total += int(cell)

        if data_type == 'float':
            return round(total, 1)
        else:
            return total

    # "Реализация лотерейных билетов"
    def sold_number_tickets(self):
        return self.counting_values_in_column(8)

    def sold_amount_tickets(self):
        return self.counting_values_in_column(9)

    def paid_number_tickets(self):
        return self.counting_values_in_column(13)

    def paid_amount_tickets(self):
        return self.counting_values_in_column(14)

    def reward_tickets(self):
        return self.counting_values_in_column(16, 'float')

    def transfer_tickets(self):
        return self.counting_values_in_column(17, 'float')

    # Проверка расчетов в каждой строке
    def check_row(self):
        min_row = int(get_boundary_values(tickets_table_range, 'min_row'))
        max_row = int(get_boundary_values(tickets_table_range, 'max_row'))

        for row in range(min_row, max_row):
            columns = {
                'sold_amount_column': 8,
                'paid_amount_column': 14,
                'percent_column': 15,
                'reward_column': 16,
                'transfer_column': 17
            }

            sold_amount = self.sheet.cell(row=row, column=columns['sold_amount_column']).value
            paid_amount = self.sheet.cell(row=row, column=columns['paid_amount_column']).value
            percent = self.sheet.cell(row=row, column=columns['percent_column']).value
            reward = float(self.sheet.cell(row=row, column=columns['reward_column']).value)
            transfer = float(self.sheet.cell(row=row, column=columns['transfer_column']).value)

            try:
                count_reward = round(sold_amount * (percent / 100), 1)
                assert count_reward == round(reward, 1)

                count_transfer = sold_amount - paid_amount - round(reward, 1)
                assert count_transfer == transfer

                return "ДА"
            except:
                return 'НЕТ!\nСтрока ' + str(row)
