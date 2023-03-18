from openpyxl import load_workbook
from util import get_cell_value, get_boundary_values


class CommonFunctions:
    def __init__(self):
        self.file_path = 'C:/Users/Trixie_LM/Desktop/1C/Отчет агента.xlsx'
        self.book = load_workbook(filename=self.file_path, data_only=True)
        self.sheet = self.book.active

    def _get_cell_value(self, column_letter, idx):
        return get_cell_value(column_letter, idx, self.sheet)


class WorkSheet(CommonFunctions):
    def realization_of_lottery_tickets(self):
        # задаем начальное и конечное значение для поиска диапазона
        start_value_row = 'Наименование филиала'
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

    def realization_of_lottery_receipts(self):
        # задаем начальное и конечное значение для поиска диапазона
        start_value_row = 'Наименование филиала'
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

    # TODO: нужен ли метод?
    def AFPS_MOSCOW_tickets(self):
        branch = "УФПС Г. МОСКВЫ"
        diapason = WorkSheet().realization_of_lottery_tickets()
        min_row = int(get_boundary_values(diapason, 'min_row'))
        max_row = int(get_boundary_values(diapason, 'max_row'))

        for row in range(min_row, max_row):
            cell = self.sheet.cell(row, 2)
            if cell.value == branch:
                print("Фраза найдена в строке:", row)
        return min_row, max_row


class ReportAgentData(CommonFunctions):
    # Беру данные из "ИТОГО" в таблице "Реализация лотерейных билетов"
    def _total_values_lottery_tickets(self, column):
        lottery_tickets_diapason = WorkSheet().realization_of_lottery_tickets()
        result_row = get_boundary_values(lottery_tickets_diapason, 'max_row')
        # Продажи
        sold_number = self._get_cell_value('F', result_row)
        sold_amount = self._get_cell_value('G', result_row)
        # Выплаты
        paid_number = self._get_cell_value('K', result_row)
        paid_amount = self._get_cell_value('L', result_row)
        # Вознаграждения
        reward = self._get_cell_value('N', result_row)
        # Перечисление принципалу
        transfer = self._get_cell_value('O', result_row)

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

    # Беру данные из "ИТОГО" в таблице "Реализация лотерейных квитанций"
    def _total_values_lottery_receipts(self, column):
        lottery_tickets_diapason = WorkSheet().realization_of_lottery_receipts()
        result_row = get_boundary_values(lottery_tickets_diapason, 'max_row')
        # Продажи
        sold_number = self._get_cell_value('C', result_row)
        sold_amount = self._get_cell_value('E', result_row)
        # Выплаты
        paid_number = self._get_cell_value('H', result_row)
        paid_amount = self._get_cell_value('J', result_row)
        # Вознаграждения
        reward = self._get_cell_value('N', result_row)
        # Перечисление принципалу
        transfer = self._get_cell_value('O', result_row)

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

    # Общая сумма двух таблиц по вознаграждению
    def _reward_of_two_tables(self):
        lottery_tickets_diapason = WorkSheet().realization_of_lottery_receipts()
        result_row = int(get_boundary_values(lottery_tickets_diapason, 'max_row')) + 5
        reward = self._get_cell_value('K', result_row)
        return reward

    # Общая сумма двух таблиц по перечислению средств
    def _transfer_of_two_tables(self):
        lottery_tickets_diapason = WorkSheet().realization_of_lottery_receipts()
        result_row = int(get_boundary_values(lottery_tickets_diapason, 'max_row')) + 6
        reward = self._get_cell_value('K', result_row)
        return reward

        # Общая сумма двух таблиц по продажам

    def _sales_of_two_tables(self):
        lottery_tickets_diapason = WorkSheet().realization_of_lottery_receipts()
        result_row = int(get_boundary_values(lottery_tickets_diapason, 'max_row')) + 3
        reward = self._get_cell_value('O', result_row)
        return reward

    # Общая сумма двух таблиц по выплатам
    def _payment_of_two_tables(self):
        lottery_tickets_diapason = WorkSheet().realization_of_lottery_receipts()
        result_row = int(get_boundary_values(lottery_tickets_diapason, 'max_row')) + 4
        reward = self._get_cell_value('O', result_row)
        return reward


class AgentAsserts(CommonFunctions):
    # Функция для подсчета итоговых данных
    # в таблицах "Реализация лотерейных билетов" и "Реализация лотерейных квитанций"
    def _counting_values_in_column(self, table, column, data_type='int'):
        total = 0
        if table == 'realization_tickets':
            diapason = WorkSheet().realization_of_lottery_tickets()
        elif table == 'realization_receipts':
            diapason = WorkSheet().realization_of_lottery_receipts()

        min_row = int(get_boundary_values(diapason, 'min_row'))
        max_row = int(get_boundary_values(diapason, 'max_row'))

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
    def _sold_number_tickets(self):
        return AgentAsserts()._counting_values_in_column('realization_tickets', 6)

    def _sold_amount_tickets(self):
        return AgentAsserts()._counting_values_in_column('realization_tickets', 7)

    def _paid_number_tickets(self):
        return AgentAsserts()._counting_values_in_column('realization_tickets', 11)

    def _paid_amount_tickets(self):
        return AgentAsserts()._counting_values_in_column('realization_tickets', 12)

    def _reward_tickets(self):
        return AgentAsserts()._counting_values_in_column('realization_tickets', 14, 'float')

    def _transfer_tickets(self):
        return AgentAsserts()._counting_values_in_column('realization_tickets', 15, 'float')

    # "Реализация лотерейных квитанций"
    def _sold_number_receipts(self):
        return AgentAsserts()._counting_values_in_column('realization_receipts', 3)

    def _sold_amount_receipts(self):
        return AgentAsserts()._counting_values_in_column('realization_receipts', 5)

    def _paid_number_receipts(self):
        return AgentAsserts()._counting_values_in_column('realization_receipts', 8)

    def _paid_amount_receipts(self):
        return AgentAsserts()._counting_values_in_column('realization_receipts', 10)

    def _reward_receipts(self):
        return AgentAsserts()._counting_values_in_column('realization_receipts', 14, 'float')

    def _transfer_receipts(self):
        return AgentAsserts()._counting_values_in_column('realization_receipts', 15, 'float')

    # Общее вознаграждение и перечисление
    def _total_rewards(self):
        tickets = AgentAsserts()._reward_tickets()
        receipts = AgentAsserts()._reward_receipts()
        return tickets + receipts

    def _total_transfer(self):
        tickets = AgentAsserts()._transfer_tickets()
        receipts = AgentAsserts()._transfer_receipts()
        return tickets + receipts

    # Проверка расчетов в каждой строке
    def _check_row(self, table):
        if table == 'realization_tickets':
            diapason = WorkSheet().realization_of_lottery_tickets()
        elif table == 'realization_receipts':
            diapason = WorkSheet().realization_of_lottery_receipts()

        min_row = int(get_boundary_values(diapason, 'min_row'))
        max_row = int(get_boundary_values(diapason, 'max_row'))

        for row in range(min_row, max_row):
            if table == 'realization_tickets':
                sold_amount_column, paid_amount_column = 7, 11
                percent_column, reward_column, transfer_column = 12, 14, 15
            elif table == 'realization_receipts':
                sold_amount_column, paid_amount_column = 5, 8
                percent_column, reward_column, transfer_column = 10, 14, 15
            sold_amount = self.sheet.cell(row=row, column=sold_amount_column).value
            paid_amount = self.sheet.cell(row=row, column=paid_amount_column).value
            percent = self.sheet.cell(row=row, column=percent_column).value
            reward = float(self.sheet.cell(row=row, column=reward_column).value)
            transfer = float(self.sheet.cell(row=row, column=transfer_column).value)

            try:
                count_reward = round(sold_amount * (percent / 100), 1)
                assert count_reward == round(reward, 1)

                count_transfer = sold_amount - paid_amount - round(reward, 1)
                assert count_transfer == transfer

                return "ДА"
            except:
                return 'НЕТ!\nСтрока ' + str(row)
