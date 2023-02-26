from openpyxl import load_workbook
from numpy import unique

file = 'C:/Users/Trixie_LM/Desktop/ExcelParser/Проданные билеты.xlsx'

book = load_workbook(filename=file, data_only=True)
sheet = book.active


class ReportSalesData:

    @staticmethod
    def total_quantity_tickets_in_report():
        return sheet['M' + str(sheet.max_row)].value

    @staticmethod
    def tickets_price_in_report():
        return sheet['N' + str(sheet.max_row)].value

    @staticmethod
    def check_cells_in_row():
        passedAssertions = 0

        for i in range(2, sheet.max_row):
            setCell = sheet['L' + str(i)].value
            drawCell = sheet['I' + str(i)].value
            ticketType = sheet['K' + str(i)].value
            lotteryType = sheet['F' + str(i)].value
            seventhDigit = sheet['J' + str(i)].value[6:7]

            try:
                if lotteryType == "Бинго лотерея" and seventhDigit == "3" and len(
                        drawCell) == 6 and setCell is not None:
                    assert ticketType == "Открытка"
                    passedAssertions += 1

                elif lotteryType == "Бинго лотерея" and seventhDigit == "3" and len(drawCell) == 6 and setCell is None:
                    assert ticketType == "Бумажный"
                    passedAssertions += 1

                elif lotteryType == "Моментальная лотерея" and setCell is None and drawCell is None:
                    assert ticketType == "Бумажный"
                    passedAssertions += 1

                elif seventhDigit == "1" and len(drawCell) == 6 and setCell is None:
                    assert ticketType == "Электронный"
                    passedAssertions += 1

                elif seventhDigit == "2" and len(drawCell) == 6 and setCell is None:
                    assert ticketType == "Купон"
                    passedAssertions += 1

                else:
                    print('Ошибка! Отчет продаж. Строка - ' + str(i))


            except AssertionError:
                print('Ошибка! Отчет продаж. Строка - ' + str(i))


        return passedAssertions


class CountTicketSales:

    @staticmethod
    def total_quantity():
        tickets_amount = 0
        for i in range(2, sheet.max_row):
            tickets_amount += sheet['M' + str(i)].value
        return tickets_amount

    @staticmethod
    def tickets_price():
        tickets_price = 0
        for i in range(2, sheet.max_row):
            tickets_price += sheet['N' + str(i)].value
        return tickets_price

    @staticmethod
    def counting_tickets(type_of_tickets, type_of_lottery=True):
        ticketsQuantity = 0
        ticketsPrice = 0
        bingoLottery = ["102021", "102041", "102051", "102091"]

        for i in range(2, sheet.max_row):
            productCode = sheet['J' + str(i)].value[0:6]
            ticketType = sheet['K' + str(i)].value
            ticketPrice = sheet['N' + str(i)].value

            # Подсчет электронных билетов и купонов
            if type_of_tickets == 'Электронный':
                if ticketType in ['Электронный', 'Купон']:
                    ticketsQuantity += 1
                    ticketsPrice += ticketPrice

            # Подсчет бумажных моменталок
            elif type_of_tickets == 'Бумажный' and type_of_lottery == 'Моментальная':
                if ticketType == type_of_tickets and productCode not in bingoLottery:
                    ticketsQuantity += 1
                    ticketsPrice += ticketPrice

            # Подсчет тиражных билетов
            elif type_of_tickets == 'Бумажный' and type_of_lottery == 'Тиражная':
                if ticketType in ['Бумажный', 'Открытка'] and productCode in bingoLottery:
                    ticketsQuantity += 1
                    ticketsPrice += ticketPrice

            # Подсчет бумажных бинго
            elif type_of_tickets == 'Бумажный':
                if ticketType == type_of_tickets:
                    ticketsQuantity += 1
                    ticketsPrice += ticketPrice

            # Подсчет отстальных (т.е. Открыток)
            else:
                if ticketType == type_of_tickets:
                    ticketsQuantity += 1
                    ticketsPrice += ticketPrice

        return ticketsQuantity, ticketsPrice

    def all_ticket_types(self):
        ticketsQuantity = self.counting_tickets('Электронный')[0] + self.counting_tickets('Бумажный')[0] + self.counting_tickets('Открытка')[0]
        ticketsPrice = self.counting_tickets('Электронный')[1] + self.counting_tickets('Бумажный')[1] + self.counting_tickets('Открытка')[1]

        return ticketsQuantity, ticketsPrice

    @staticmethod
    def unique_ticket_numbers():
        ticketNumbersQuantity = 0
        ticketNumbers = []
        for i in range(2, sheet.max_row):
            ticketNumber = sheet['J' + str(i)].value
            if ticketNumber is not None:
                ticketNumbers.append(ticketNumber)
                ticketNumbersQuantity += 1

        if ticketNumbersQuantity == len(unique(ticketNumbers)):
            return "Все билеты в отчете уникальные"
        else:
            return "В отчете есть дубликат билета"

    @staticmethod
    def tickets_of_sets_in_report():
        setAmount = 0
        allSets = []
        for i in range(2, sheet.max_row):

            ticketSetNumber = sheet['L' + str(i)].value
            if ticketSetNumber is not None:
                allSets.append(ticketSetNumber)
                setAmount += 1

        if setAmount % 3 == 0 and setAmount / 3 == len(unique(allSets)):
            return "Все билеты набора находятся в отчете"
        else:
            return "Нужно проверить открытки"

import report_of_checking
print(  ReportSalesData().total_quantity_tickets_in_report(), float(report_of_checking.total_number_of_sales()))