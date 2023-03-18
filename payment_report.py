from openpyxl import load_workbook
from numpy import unique

file = 'C:/Users/Trixie_LM/Desktop/1C/Выплаченные выигрыши.xlsx'

book = load_workbook(filename=file, data_only=True)
sheet = book.active


class ReportPaymentsData:

    @staticmethod
    def total_quantity_tickets_in_report():
        return sheet['M' + str(sheet.max_row)].value

    @staticmethod
    def win_amount_in_report():
        return sheet['N' + str(sheet.max_row)].value

    @staticmethod
    def check_cells_in_row():
        passedAssertions = 0
        bingoLottery = ["102021", "102041", "102051", "102091"]

        for i in range(2, sheet.max_row):
            setCell = sheet['K' + str(i)].value
            drawCell = sheet['H' + str(i)].value
            ticketType = sheet['J' + str(i)].value
            seventhDigit = sheet['I' + str(i)].value[6:7]
            productCode = sheet['I' + str(i)].value[0:6]
            paymentType = sheet['L' + str(i)].value

            try:
                if productCode in bingoLottery and seventhDigit == "3" and len(drawCell) == 6 and setCell is not None:
                    assert ticketType == "Открытка" and paymentType == "По билету"
                    passedAssertions += 1

                elif productCode in bingoLottery and seventhDigit == "3" and len(drawCell) == 6 and setCell is None:
                    assert ticketType == "Бумажный" and paymentType == "По билету"
                    passedAssertions += 1

                elif productCode not in bingoLottery and setCell is None and drawCell is None:
                    assert ticketType == "Бумажный" and paymentType == "По билету"
                    passedAssertions += 1

                elif seventhDigit in ["1", "2"] and len(drawCell) == 6 and setCell is None:
                    assert ticketType in ["Электронный", "Купон"]
                    assert paymentType in ["По билету", "По билету с проверочными кодом", "По номеру телефона"]
                    passedAssertions += 1

                else:
                    print(
                        f'Ошибка! Отчет выплат. Строка - {str(i)}\n{sheet["I" + str(i)].value}-{len(drawCell)}{setCell}')

            except AssertionError:
                print('Ошибка! Отчет выплат. Строка - ' + str(i))

        return passedAssertions


class PaymentsAsserts:

    @staticmethod
    def total_quantity():
        tickets_amount = 0
        for i in range(2, sheet.max_row):
            tickets_amount += sheet['M' + str(i)].value
        return tickets_amount

    @staticmethod
    def win_amount():
        win_amount = 0
        for i in range(2, sheet.max_row):
            win_amount += sheet['N' + str(i)].value
        return win_amount

    @staticmethod
    def counting_tickets(type_of_tickets, type_of_lottery=True):
        ticketsQuantity = 0
        winAmount = 0
        bingoLottery = ["102021", "102041", "102051", "102091"]

        for i in range(2, sheet.max_row):
            productCode = sheet['I' + str(i)].value[0:6]
            ticketType = sheet['J' + str(i)].value
            winAmountCell = sheet['N' + str(i)].value

            # Подсчет электронных билетов и купонов
            if type_of_tickets == 'Электронный':
                if ticketType in ['Электронный', 'Купон']:
                    ticketsQuantity += 1
                    winAmount += winAmountCell

            # Подсчет бумажных моменталок
            elif type_of_tickets == 'Бумажный' and type_of_lottery == 'Моментальная':
                if ticketType == type_of_tickets and productCode not in bingoLottery:
                    ticketsQuantity += 1
                    winAmount += winAmountCell

            # Подсчет тиражных билетов
            elif type_of_tickets == 'Бумажный' and type_of_lottery == 'Тиражная':
                if ticketType in ['Бумажный', 'Открытка'] and productCode in bingoLottery:
                    ticketsQuantity += 1
                    winAmount += winAmountCell

            # Подсчет бумажных
            elif type_of_tickets == 'Бумажный':
                if ticketType == type_of_tickets:
                    ticketsQuantity += 1
                    winAmount += winAmountCell

            # Подсчет отстальных (т.е. Открыток)
            else:
                if ticketType == type_of_tickets:
                    ticketsQuantity += 1
                    winAmount += winAmountCell

        return ticketsQuantity, winAmount

    def all_ticket_types(self):
        ticketsQuantity = self.counting_tickets('Электронный')[0] + self.counting_tickets('Бумажный')[0] + \
                          self.counting_tickets('Открытка')[0]
        winAmounts = self.counting_tickets('Электронный')[1] + self.counting_tickets('Бумажный')[1] + \
                     self.counting_tickets('Открытка')[1]

        return ticketsQuantity, winAmounts

    @staticmethod
    def unique_ticket_numbers():
        ticketNumbersQuantity = 0
        ticketNumbers = []
        for i in range(2, sheet.max_row):
            ticketNumber = sheet['I' + str(i)].value
            if ticketNumber is not None:
                ticketNumbers.append(ticketNumber)
                ticketNumbersQuantity += 1

        if ticketNumbersQuantity == len(unique(ticketNumbers)):
            return "Все билеты в отчете уникальные"
        else:
            return "В отчете есть дубликат билета"

    @staticmethod
    def win_amount_less_15000():
        winAmountMore15000 = 0

        for i in range(2, sheet.max_row):
            ticketSetNumber = sheet['N' + str(i)].value
            if ticketSetNumber >= 15000:
                winAmountMore15000 += 1

        if winAmountMore15000 == 0:
            return "Отсутствуют билеты с выигрышем более 15000 руб"
        else:
            return "В отчете есть билеты с выигрышем более 15000 руб"

    # Необходим для paid_payments_registry
    @staticmethod
    def collecting_numbers():
        arrayNumbers = {}

        for i in range(2, sheet.max_row):
            ticketNumber = sheet['I' + str(i)].value
            winAmount = int(sheet['N' + str(i)].value)

            arrayNumbers[ticketNumber] = winAmount

        return arrayNumbers

