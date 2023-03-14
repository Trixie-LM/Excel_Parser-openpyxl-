from openpyxl import load_workbook
from sales_report import ReportSalesData, CountTicketSales
from payment_report import ReportPaymentsData, CountTicketPayments
import report_of_checking


def editing_cells():
    array = [
        # Отчет по продажам
        "B2:F3", "B4:E5", "B6:E7", "B8:E9", "B10:E11", "B12:E13", "B14:E15", "F4", "F5", "F6", "F7", "F8", "F9",
        "F10",
        "F11", "F12", "F13", "F14", "F15",
        "B16:F16",  # РАЗДЕЛИТЕЛЬ
        "B17:E17", "B18:E18", "B19:E20", "B21:E22", "F21", "F22", "B23:F23", "B24:F24",
        "F17", "F18", "F19:F20",

        # Отчет для сверки ф.130 продаж
        "H2:I3", "H4:I4", "H5:I5", "H6:I6", "H7:I7",
        "H8:I11",  # РАЗДЕЛИТЕЛЬ
        "H12:I12", "H13:I13", "H14:I14", "H15:I15",

        # Отчет по выплатам
        "L2:P3", "L4:O5", "L6:O7", "L8:O9", "L10:O11", "L12:O13", "L14:O15", "P4", "P5", "P6", "P7", "P8", "P9",
        "P10",
        "P11", "P12", "P13", "P14", "P15",
        "L16:P16",  # РАЗДЕЛИТЕЛЬ
        "L17:O17", "L18:O18", "L19:O20", "L21:O22", "P21", "P22", "L23:P23", "L24:P24",
        "P17", "P18", "P19:P20",

        # Отчет для сверки ф.130 выплат
        "R2:S3", "R4:S4", "R5:S5", "R6:S6", "R7:S7",
        "R8:S11",  # РАЗДЕЛИТЕЛЬ
        "R12:S12", "R13:S13", "R14:S14", "R15:S15"

    ]
    return array

#TODO
# Внедрить метод в проект
# Разделение элементов для input_data()
def startCell(number):
    return editing_cells()[number].split(':')[0]

# Импорт данных в таблицу
def input_data():
    array = [
        # Отчет по продажам
        ("B2", "Отчет по продажам"),
        ("B4", "Итоговое количество проданных билетов в отчете"),
        ("F4", f"{ReportSalesData().total_quantity_tickets_in_report()} шт"),
        ("F5", f"{ReportSalesData().tickets_price_in_report()} руб"),
        ("B6", "Электронные билеты"),
        ("F6", f"{CountTicketSales().counting_tickets('Электронный')[0]} шт"),
        ("F7", f"{CountTicketSales().counting_tickets('Электронный')[1]} руб"),
        ("B8", "Бумажные билеты"),
        ("F8", f"{CountTicketSales().counting_tickets('Бумажный')[0]} шт"),
        ("F9", f"{CountTicketSales().counting_tickets('Бумажный')[1]} руб"),
        ("B10", "Наборы/Открытки"),
        ("F10", f"{CountTicketSales().counting_tickets('Открытка')[0]} шт"),
        ("F11", f"{CountTicketSales().counting_tickets('Открытка')[1]} руб"),
        ("B12", "Тиражные билеты"),
        ("F12", f"{CountTicketSales().counting_tickets('Бумажный', 'Тиражная')[0]} шт"),
        ("F13", f"{CountTicketSales().counting_tickets('Бумажный', 'Тиражная')[1]} руб"),
        ("B14", "Билеты моментальной лотереи"),
        ("F14", f"{CountTicketSales().counting_tickets('Бумажный', 'Моментальная')[0]} шт"),
        ("F15", f"{CountTicketSales().counting_tickets('Бумажный', 'Моментальная')[1]} руб"),
        # РАЗДЕЛИТЕЛЬ
        ("B17", "Сложение билетов по столбцу \"Кол-во\""), ("F17", f"{CountTicketSales().total_quantity()} шт"),
        ("B18", "Проверка точности атрибутов для билета"), ("F18", f"{ReportSalesData().check_cells_in_row()} шт"),
        ("B19", "Сложение цены по столбцу \n\"Стоимость билета\""),
        ("F19", f"{CountTicketSales().tickets_price()} руб"),
        ("B21", f"Общая сумма ячеек \"Электронные билеты\", \"Бумажные билеты\" и \"Открытки\" "),
        ("F21", f"{CountTicketSales().all_ticket_types()[0]} шт "),
        ("F22", f"{CountTicketSales().all_ticket_types()[1]} руб"),
        ("B23", CountTicketSales().unique_ticket_numbers()),
        ("B24", CountTicketSales().tickets_of_sets_in_report()),

        # Отчет для сверки ф.130 продаж
        ("H2", "Отчет для сверки \n продаж ф.130"),
        ("H4", f"{report_of_checking.total_number_of_sales()} шт"),
        ("H5", f"{report_of_checking.total_amount_of_sales()} руб"),
        ("H6", f"{report_of_checking.number_of_digital_sales()} шт"),
        ("H7", f"{report_of_checking.amount_of_digital_sales()} руб"),
        ("H12", f"{report_of_checking.number_of_circulation_sales()} шт"),
        ("H13", f"{report_of_checking.amount_of_circulation_sales()} руб"),
        ("H14", f"{report_of_checking.number_of_instant_sales()} шт"),
        ("H15", f"{report_of_checking.amount_of_instant_sales()} руб"),

        # Отчет по выплатам
        ("L2", "Отчет по выплатам"),
        ("L4", "Итоговое количество выплаченных билетов в отчете"),
        ("P4", f"{ReportPaymentsData().total_quantity_tickets_in_report()} шт"),
        ("P5", f"{ReportPaymentsData().win_amount_in_report()} руб"),
        ("L6", "Электронные билеты"),
        ("P6", f"{CountTicketPayments().counting_tickets('Электронный')[0]} шт"),
        ("P7", f"{CountTicketPayments().counting_tickets('Электронный')[1]} руб"),
        ("L8", "Бумажные билеты"),
        ("P8", f"{CountTicketPayments().counting_tickets('Бумажный')[0]} шт"),
        ("P9", f"{CountTicketPayments().counting_tickets('Бумажный')[1]} руб"),
        ("L10", "Наборы/Открытки"),
        ("P10", f"{CountTicketPayments().counting_tickets('Открытка')[0]} шт"),
        ("P11", f"{CountTicketPayments().counting_tickets('Открытка')[1]} руб"),
        ("L12", "Тиражные билеты"),
        ("P12", f"{CountTicketPayments().counting_tickets('Бумажный', 'Тиражная')[0]} шт"),
        ("P13", f"{CountTicketPayments().counting_tickets('Бумажный', 'Тиражная')[1]} руб"),
        ("L14", "Билеты моментальной лотереи"),
        ("P14", f"{CountTicketPayments().counting_tickets('Бумажный', 'Моментальная')[0]} шт"),
        ("P15", f"{CountTicketPayments().counting_tickets('Бумажный', 'Моментальная')[1]} руб"),
        # РАЗДЕЛИТЕЛЬ
        ("L17", "Сложение билетов по столбцу \"Кол-во\""), ("P17", f"{CountTicketPayments().total_quantity()} шт"),
        ("L18", "Проверка точности атрибутов для билета"),
        ("P18", f"{ReportPaymentsData().check_cells_in_row()} шт"),
        ("L19", "Сложение по столбцу \n\"Размер выигрыша\""), ("P19", f"{CountTicketPayments.win_amount()} руб"),
        ("L21", f"Общая сумма ячеек \"Электронные билеты\", \"Бумажные билеты\" и \"Открытки\""),
        ("P21", f"{CountTicketPayments().all_ticket_types()[0]} шт"),
        ("P22", f"{CountTicketPayments().all_ticket_types()[1]} руб"),
        ("L23", CountTicketPayments().unique_ticket_numbers()),
        ("L24", CountTicketPayments().win_amount_less_15000()),

        # Отчет для сверки ф.130 выплат
        ("R2", "Отчет для сверки \n выплат ф.130"),
        ("R4", f'{report_of_checking.total_number_of_payments()} шт'),
        ("R5", f'{report_of_checking.total_amount_of_payments()} руб'),
        ("R6", f'{report_of_checking.number_of_digital_payments()} шт'),
        ("R7", f'{report_of_checking.amount_of_digital_payments()} руб'),
        ("R12", f'{report_of_checking.number_of_circulation_payments()} шт'),
        ("R13", f'{report_of_checking.amount_of_circulation_payments()} руб'),
        ("R14", f'{report_of_checking.number_of_instant_payments()} шт'),
        ("R15", f'{report_of_checking.amount_of_instant_payments()} руб')
    ]
    return array


# Сверка данных и заливка фона ячейки
def check_and_painting():
    array = [
        # Отчет по продажам
        ("F17", ReportSalesData().total_quantity_tickets_in_report(), CountTicketSales().total_quantity()),
        (
            "F18", ReportSalesData().total_quantity_tickets_in_report(),
            float(ReportSalesData().check_cells_in_row())),
        ("F21", ReportSalesData().total_quantity_tickets_in_report(),
         float(CountTicketSales().all_ticket_types()[0])),
        ("H4", ReportSalesData().total_quantity_tickets_in_report(),
         float(report_of_checking.total_number_of_sales())),

        ("F19", ReportSalesData().tickets_price_in_report(), CountTicketSales().tickets_price()),
        ("F22", ReportSalesData().tickets_price_in_report(), CountTicketSales().all_ticket_types()[1]),
        ("H5", ReportSalesData().tickets_price_in_report(), float(report_of_checking.total_amount_of_sales())),

        ("H6", CountTicketSales().counting_tickets('Электронный')[0], report_of_checking.number_of_digital_sales()),
        ("H7", CountTicketSales().counting_tickets('Электронный')[1],
         float(report_of_checking.amount_of_digital_sales())),
        ("H12", CountTicketSales().counting_tickets('Бумажный', 'Тиражная')[0],
         report_of_checking.number_of_circulation_sales()),
        ("H13", CountTicketSales().counting_tickets('Бумажный', 'Тиражная')[1],
         float(report_of_checking.amount_of_circulation_sales())),
        ("H14", CountTicketSales().counting_tickets('Бумажный', 'Моментальная')[0],
         report_of_checking.number_of_instant_sales()),
        ("H15", CountTicketSales().counting_tickets('Бумажный', 'Моментальная')[1],
         float(report_of_checking.amount_of_instant_sales())),

        ("B23", CountTicketSales().unique_ticket_numbers(), 'Все билеты в отчете уникальные'),
        ("B24", CountTicketSales().tickets_of_sets_in_report(), 'Все билеты набора находятся в отчете'),

        # Отчет по выплатам
        ("P17", ReportPaymentsData().total_quantity_tickets_in_report(), CountTicketPayments().total_quantity()),
        ("P18", ReportPaymentsData().total_quantity_tickets_in_report(),
         float(ReportPaymentsData().check_cells_in_row())),
        (
            "P21", ReportPaymentsData().total_quantity_tickets_in_report(),
            float(CountTicketPayments().all_ticket_types()[0])),
        ("R4", ReportPaymentsData().total_quantity_tickets_in_report(),
         float(report_of_checking.total_number_of_payments())),

        ("P19", ReportPaymentsData().win_amount_in_report(), CountTicketPayments().win_amount()),
        ("P22", ReportPaymentsData().win_amount_in_report(), CountTicketPayments().all_ticket_types()[1]),
        ("R5", ReportPaymentsData().win_amount_in_report(), float(report_of_checking.total_amount_of_payments())),

        ("R6", CountTicketPayments().counting_tickets('Электронный')[0],
         report_of_checking.number_of_digital_payments()),
        ("R7", CountTicketPayments().counting_tickets('Электронный')[1],
         float(report_of_checking.amount_of_digital_payments())),
        ("R12", CountTicketPayments().counting_tickets('Бумажный', 'Тиражная')[0],
         report_of_checking.number_of_circulation_payments()),
        ("R13", CountTicketPayments().counting_tickets('Бумажный', 'Тиражная')[1],
         float(report_of_checking.amount_of_circulation_payments())),
        ("R14", CountTicketPayments().counting_tickets('Бумажный', 'Моментальная')[0],
         report_of_checking.number_of_instant_payments()),
        ("R15", CountTicketPayments().counting_tickets('Бумажный', 'Моментальная')[1],
         float(report_of_checking.amount_of_instant_payments())),

        ("L23", CountTicketPayments().unique_ticket_numbers(), 'Все билеты в отчете уникальные'),
        ("L24", CountTicketPayments().win_amount_less_15000(), 'Отсутствуют билеты с выигрышем более 15000 руб')
    ]
    return array


