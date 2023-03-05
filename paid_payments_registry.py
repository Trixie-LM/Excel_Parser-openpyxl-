from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from payment_report import CountTicketPayments

file = 'C:/Users/Trixie_LM/Desktop/1C/Реестр выплаченных выигрышей.xlsx'
book = load_workbook(filename=file, data_only=True)
sheet = book.active


def total_quantity_tickets_in_report():
    max_column = int(sheet.max_column) - 3


    for number in range(0, sheet.max_row):
        print(get_column_letter(number))

    #     if sheet.max_column:
    #         ticket_numbers_column = get_column_letter(number)
    #         win_amounts_column = get_column_letter(number + 1)
    #
    # return sheet['M' + str(sheet.max_row)].value


# ttl = int(sheet.max_column) - 3
# for number in range(ttl, 1, -4):
#     qqq = get_column_letter(number)

from openpyxl import load_workbook as lw
from openpyxl.utils import get_column_letter



for col in range(int(sheet.max_column-4) - 3, 8, -4):
    col_letter = get_column_letter(col)
    in_total = [cell for cell in sheet[col_letter] if cell.value == 'Итого по каждой игре'][0].row
    qqq = [cell for cell in sheet[col_letter][in_total:] if cell.value == 'ИТОГО:'][0].row






    print(qqq)









def win_amount_in_report():
    return sheet['N' + str(sheet.max_row)].value


# Поиск расхождений между 2 отчетами
def discrepancy_reports():
    remaining_tickets = CountTicketPayments.collecting_numbers()
    try:
        for number in range(3, int(sheet.max_column), 4):
            ticket_numbers_column = get_column_letter(number)
            win_amounts_column = get_column_letter(number + 1)

            for row in range(1, sheet.max_row + 1):
                ticket_number_cell = sheet[ticket_numbers_column + str(row)].value
                win_amount_cell = sheet[win_amounts_column + str(row)].value

                if ticket_number_cell is not None and type(
                        ticket_number_cell) != float and ticket_number_cell.isdigit():
                    if remaining_tickets[ticket_number_cell] == int(win_amount_cell.replace(' ', '')):
                        remaining_tickets.pop(ticket_number_cell)
    except ValueError:
        remaining_tickets = f"Что-то не так с билетом - {ticket_number_cell}, возможно, нет в отчете выплат!"

    return remaining_tickets


# Счет билетов по столбцу "Номер лотерейного билета..."
def count_tickets_and_winnings():
    amount_tickets = 0
    amount_winnings = 0

    for number in range(3, int(sheet.max_column), 4):
        ticket_numbers_column = get_column_letter(number)
        win_amounts_column = get_column_letter(number + 1)

        for row in range(1, sheet.max_row + 1):
            ticket_number_cell = sheet[ticket_numbers_column + str(row)].value
            win_amount_column = sheet[win_amounts_column + str(row)].value

            if ticket_number_cell is not None and type(ticket_number_cell) != float and ticket_number_cell.isdigit():
                amount_tickets += 1
                amount_winnings += int(win_amount_column.replace(' ', ''))

    return amount_tickets, amount_winnings



