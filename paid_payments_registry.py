from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from payment_report import PaymentsAsserts

file = 'C:/Users/Trixie_LM/Desktop/1C/Реестр выплаченных выигрышей.xlsx'
book = load_workbook(filename=file)
sheet = book.active

# Проверка наличия доп.листа в файле
if "Реестр выплаченных выигрышей" not in book.sheetnames:
    book.create_sheet("Реестр выплаченных выигрышей")

add_sheet = book["Реестр выплаченных выигрышей"]


class PreCondition:
    @staticmethod
    def copying_table():
        add_sheet.delete_rows(1, add_sheet.max_row)
        # Копирование основной части файла в итоговый отчет
        begin_column = int(sheet.max_column) - 3
        second_last_page = sheet.iter_rows(min_col=begin_column - 4, max_col=begin_column - 1, values_only=True)
        last_page = sheet.iter_rows(min_col=begin_column, values_only=True)

        # Копирование строк в новый файл
        for row in second_last_page:
            add_sheet.append(row)
        for row in last_page:
            add_sheet.append(row)

        book.save(file)

    @staticmethod
    def delete_rows():
        # Удаление лишних строк в начале
        first_row = 0
        for row in list(add_sheet.rows):
            for cell in row:
                if cell.value == "Итого по каждой игре":
                    first_row = cell.row
                    break
            if first_row > 0:
                break

        if first_row > 0:
            add_sheet.delete_rows(1, first_row)

        # Удаление лишних строк в конце
        last_row = 0
        for row in reversed(list(add_sheet.rows)):
            for cell in row:
                if cell.value == "ИТОГО:":
                    last_row = cell.row
                    break
            if last_row > 0:
                break

        if last_row > 0:
            add_sheet.delete_rows(last_row + 1, add_sheet.max_row - last_row)

        book.save(file)


def paid_lottery_names_list():
    return list(add_sheet.rows)

# TODO: добавить в отчет для проверки
# Поиск расхождение между реестром и отчетом по выплатам
def discrepancy_reports():
    remaining_tickets = PaymentsAsserts.collecting_numbers()
    try:
        for number in range(3, int(sheet.max_column), 4):
            ticket_numbers_column = get_column_letter(number)
            win_amounts_column = get_column_letter(number + 1)

            for row in range(1, sheet.max_row + 1):
                ticket_number_cell = sheet[ticket_numbers_column + str(row)].value
                win_amount_cell = sheet[win_amounts_column + str(row)].value

                if ticket_number_cell is not None and type(
                        ticket_number_cell) != float and ticket_number_cell.isdigit():
                    print(remaining_tickets[ticket_number_cell])
                    if remaining_tickets[ticket_number_cell] == int(win_amount_cell.replace(' ', '')):
                        remaining_tickets.pop(ticket_number_cell)
    except ValueError:
        remaining_tickets = f"Ошибка! Что-то не так с билетом \"{ticket_number_cell}\" на строке {row}, возможно, нет в отчете выплат!"

    except KeyError:
        remaining_tickets = f"Ошибка! Билета \"{ticket_number_cell}\" на строке {row} нет в реестре выплаченных выигрышей, но есть в отчете по продажам!"

    return remaining_tickets

# TODO: добавить в отчет для проверки
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

            if ticket_number_cell is not None and type(ticket_number_cell) != int and ticket_number_cell.isdigit():
                amount_tickets += 1
                amount_winnings += int(win_amount_column.replace(' ', ''))

    return amount_tickets, amount_winnings
