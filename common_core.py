from openpyxl import load_workbook
from sales_report import ReportSalesData, CountTicketSales
from payment_report import ReportPaymentsData, CountTicketPayments
from branch_report import ReportBranchData, BranchAsserts
from agent_report import ReportAgentData, AgentAsserts
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
        "R12:S12", "R13:S13", "R14:S14", "R15:S15",

        # Реестр выплаченных выигрышей
        "U2:X2",
        "U3:V3", "W3:W3", "X3:X3",
        "U4:V4", "W4:W4", "X4:X4",
        "U5:V5", "W5:W5", "X5:X5",
        "U6:V6", "W6:W6", "X6:X6",
        "U7:V7", "W7:W7", "X7:X7",
        "U8:V8", "W8:W8", "X8:X8",
        "U9:V9", "W9:W9", "X9:X9",
        "U10:V10", "W10:W10", "X10:X10",
        "U11:V11", "W11:W11", "X11:X11",
        "U12:V12", "W12:W12", "X12:X12",
        "U13:V13", "W13:W13", "X13:X13",
        "U14:V14", "W14:W14", "X14:X14",

        # Отчет филиала
        "B28:G28", "B29:D29", "E29:G29", "B30:D31", "E30:G31",
        "B32:C32", "B33:C33", "D32:E33", "F32:G32", "F33:G33",
        "B34:C34", "B35:C35", "D34:E35", "F34:G34", "F35:G35",
        "B36:C37", "D36:E37", "F36:G37",
        "B38:C39", "D38:E39", "F38:G39",
        "B40:G40",  # РАЗДЕЛИТЕЛЬ
        "B41:C41", "B42:C42", "D41:E42", "F41:G41", "F42:G42",
        "B43:C43", "B44:C44", "D43:E44", "F43:G43", "F44:G44",
        "B45:C46", "D45:E46", "F45:G46",
        "B47:C48", "D47:E48", "F47:G48",
        "B49:B50", "C49:F50", "G49:G50",
        "B51:G52", "B53:G54",

        # Отчет агента
        "J28:O28", "J29:L29", "M29:O29", "J30:L31", "M30:O31",
        "J32:K32", "J33:K33", "L32:M33", "N32:O32", "N33:O33",
        "J34:K34", "J35:K35", "L34:M35", "N34:O34", "N35:O35",
        "J36:K37", "L36:M37", "N36:O37",
        "J38:K39", "L38:M39", "N38:O39",
        "J40:O40",  # РАЗДЕЛИТЕЛЬ
        "J41:K41", "J42:K42", "L41:M42", "N41:O41", "N42:O42",
        "J43:K43", "J44:K44", "L43:M44", "N43:O43", "N44:O44",
        "J45:K46", "L45:M46", "N45:O46",
        "J47:K48", "L47:M48", "N47:O48",
        "J49:J50", "K49:N50", "O49:O50",
        "J51:O52", "J53:O54",

        # Отчет агента для бестиражных лотерей
        "R28:U28",
        "R29:S29", "R30:S30", "T29:U30",
        "R31:S31", "R32:S32", "T31:U32",
        "R33:S34", "T33:U34", "R35:S36", "T35:U36",
        "R37:U37",  # РАЗДЕЛИТЕЛЬ
        "R38:S38", "R39:S39", "T38:U39",
        "R40:S40", "R41:S41", "T40:U41",
        "R42:S43", "T42:U43", "R44:S45", "T44:U45",
        "R46:R47", "S46:U47",
        "R48:U49", "R50:U51"
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
        ("R15", f'{report_of_checking.amount_of_instant_payments()} руб'),

        # Реестр выплаченных выигрышей
        ("U2", "Реестр выплаченных выигрышей"),

        # Отчет филиала
        ("B28", "Отчет филиала"),
        ("B29", "Билеты (1 табл)"), ("E29", "Квитанции (2 табл)"),
        ("B30", "Сумма всех проданных\nтиражно-бумажных и открыток"), ("E30", "Сумма всех проданных\nэлектронных и купонов"),
        ("D32", "Продажа"),
        ("B32", f"{ReportBranchData()._total_values_lottery_tickets('sold_number')} шт"),
        ("B33", f"{ReportBranchData()._total_values_lottery_tickets('sold_amount')} руб"),
        ("F32", f"{ReportBranchData()._total_values_lottery_receipts('sold_number')} шт"),
        ("F33", f"{ReportBranchData()._total_values_lottery_receipts('sold_amount')} руб"),
        ("D34", "Выплата"),
        ("B34", f"{ReportBranchData()._total_values_lottery_tickets('paid_number')} шт"),
        ("B35", f"{ReportBranchData()._total_values_lottery_tickets('paid_amount')} руб"),
        ("F34", f"{ReportBranchData()._total_values_lottery_receipts('paid_number')} шт"),
        ("F35", f"{ReportBranchData()._total_values_lottery_receipts('paid_amount')} руб"),
        ("D36", "Вознаграждение Филиала"),
        ("B36", f"{ReportBranchData()._total_values_lottery_tickets('reward')} руб"),
        ("F36", f"{ReportBranchData()._total_values_lottery_receipts('reward')} руб"),
        ("D38", "Перечислению за отчетный период"),
        ("B38", f"{ReportBranchData()._total_values_lottery_tickets('transfer')} руб"),
        ("F38", f"{ReportBranchData()._total_values_lottery_receipts('transfer')} руб"),
        # РАЗДЕЛИТЕЛЬ
        ("D41", "Подсчет по столбцу \"Реализовано бил..\""),
        ("B41", f"{BranchAsserts()._sold_number_tickets()} шт"), ("F41", f"{BranchAsserts()._sold_number_receipts()} шт"),
        ("B42", f"{BranchAsserts()._sold_amount_tickets()} руб"), ("F42", f"{BranchAsserts()._sold_amount_receipts()} руб"),
        ("D43", "Подсчет по столбцу \"Выплачено выигр..\""),
        ("B43", f"{BranchAsserts()._paid_number_tickets()} шт"), ("F43", f"{BranchAsserts()._paid_number_receipts()} шт"),
        ("B44", f"{BranchAsserts()._paid_amount_tickets()} руб"), ("F44", f"{BranchAsserts()._paid_amount_receipts()} руб"),
        ("D45", "Подсчет по столбцу \"Вознаграждение Ф..\""),
        ("B45", f"{BranchAsserts()._reward_tickets()} руб"), ("F45", f"{BranchAsserts()._reward_receipts()} руб"),
        ("D47", "Подсчет по столбцу \"Подлежит перечис..\""),
        ("B47", f"{BranchAsserts()._transfer_tickets()} руб"), ("F47", f"{BranchAsserts()._transfer_receipts()} руб"),
        ("C49", "Расчет для каждой строки верный?\nЕсли нет, то на какой строке?"),
        ("B49", f"{BranchAsserts()._check_row('realization_tickets')}"), ("G49", f"{BranchAsserts()._check_row('realization_receipts')}"),
        ("B51", f"Общее вознаграждение филиала составило:\n{ReportBranchData()._reward_of_two_tables()} руб"),
        ("B53", f"Общая сумма к перечислению на расч.счет составляет:\n{ReportBranchData()._transfer_of_two_tables()} руб"),

        # Отчет агента
        ("J28", "Отчет агента"),
        ("J29", "Билеты (1 табл)"), ("M29", "Квитанции (2 табл)"),
        ("J30", "Сумма всех проданных\nтиражно-бумажных и открыток"),
        ("M30", "Сумма всех проданных\nэлектронных и купонов"),
        ("L32", "Продажа"),
        ("J32", f"{ReportAgentData()._total_values_lottery_tickets('sold_number')} шт"),
        ("J33", f"{ReportAgentData()._total_values_lottery_tickets('sold_amount')} руб"),
        ("N32", f"{ReportAgentData()._total_values_lottery_receipts('sold_number')} шт"),
        ("N33", f"{ReportAgentData()._total_values_lottery_receipts('sold_amount')} руб"),
        ("L34", "Выплата"),
        ("J34", f"{ReportAgentData()._total_values_lottery_tickets('paid_number')} шт"),
        ("J35", f"{ReportAgentData()._total_values_lottery_tickets('paid_amount')} руб"),
        ("N34", f"{ReportAgentData()._total_values_lottery_receipts('paid_number')} шт"),
        ("N35", f"{ReportAgentData()._total_values_lottery_receipts('paid_amount')} руб"),
        ("L36", "Вознаграждение Филиала"),
        ("J36", f"{ReportAgentData()._total_values_lottery_tickets('reward')} руб"),
        ("N36", f"{ReportAgentData()._total_values_lottery_receipts('reward')} руб"),
        ("L38", "Перечислению за отчетный период"),
        ("J38", f"{ReportAgentData()._total_values_lottery_tickets('transfer')} руб"),
        ("N38", f"{ReportAgentData()._total_values_lottery_receipts('transfer')} руб"),
            # РАЗДЕЛИТЕЛЬ
        ("L41", "Подсчет по столбцу \"Реализовано бил..\""),
        ("J41", f"{AgentAsserts()._sold_number_tickets()} шт"),
        ("N41", f"{AgentAsserts()._sold_number_receipts()} шт"),
        ("J42", f"{AgentAsserts()._sold_amount_tickets()} руб"),
        ("N42", f"{AgentAsserts()._sold_amount_receipts()} руб"),
        ("L43", "Подсчет по столбцу \"Выплачено выигр..\""),
        ("J43", f"{AgentAsserts()._paid_number_tickets()} шт"),
        ("N43", f"{AgentAsserts()._paid_number_receipts()} шт"),
        ("J44", f"{AgentAsserts()._paid_amount_tickets()} руб"),
        ("N44", f"{AgentAsserts()._paid_amount_receipts()} руб"),
        ("L45", "Подсчет по столбцу \"Вознаграждение Ф..\""),
        ("J45", f"{AgentAsserts()._reward_tickets()} руб"), ("N45", f"{AgentAsserts()._reward_receipts()} руб"),
        ("L47", "Подсчет по столбцу \"Подлежит перечис..\""),
        ("J47", f"{AgentAsserts()._transfer_tickets()} руб"), ("N47", f"{AgentAsserts()._transfer_receipts()} руб"),
        ("K49", "Расчет для каждой строки верный?\nЕсли нет, то на какой строке?"),
        ("J49", f"{AgentAsserts()._check_row('realization_tickets')}"),
        ("O49", f"{AgentAsserts()._check_row('realization_receipts')}"),
        ("J51", f"Общее вознаграждение агента составило:\n{ReportAgentData()._reward_of_two_tables()} руб"),
        ("J53",
         f"Общая сумма к перечислению на расч.счет составляет:\n{ReportAgentData()._transfer_of_two_tables()} руб"),

        # Отчет агента для бестиражных лотерей
        ("R28", "Отчет агента для бестиражных лотерей"),
        ("T29", "Продажа"),
        ("T31", "Выплата"),
        ("T33", "Вознаграждение Филиала"),
        ("T35", "Перечислению за отчетный период"),
            # РАЗДЕЛИТЕЛЬ
        ("T38", "Подсчет по столбцу \"Реализовано бил..\""),
        ("T40", "Подсчет по столбцу \"Выплачено выигр..\""),
        ("T42", "Подсчет по столбцу \"Вознаграждение Ф..\""),
        ("T44", "Подсчет по столбцу \"Подлежит перечис..\""),
        ("S46", "Расчет для каждой строки верный?\nЕсли нет, то на какой строке?"),
        ("R48", "Вознаграждение филиала составило:\n xxx"),
        ("R50", "К перечислению на расч.счет следует:\n xxx")
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
        ("L24", CountTicketPayments().win_amount_less_15000(), 'Отсутствуют билеты с выигрышем более 15000 руб'),

        # Отчет филиала
        ("B41", BranchAsserts()._sold_number_tickets(), ReportBranchData()._total_values_lottery_tickets('sold_number')),
        ("F41", BranchAsserts()._sold_number_receipts(), ReportBranchData()._total_values_lottery_receipts('sold_number')),
        ("B42", BranchAsserts()._sold_amount_tickets(), ReportBranchData()._total_values_lottery_tickets('sold_amount')),
        ("F42", BranchAsserts()._sold_amount_receipts(), ReportBranchData()._total_values_lottery_receipts('sold_amount')),
        ("B43", BranchAsserts()._paid_number_tickets(), ReportBranchData()._total_values_lottery_tickets('paid_number')),
        ("F43", BranchAsserts()._paid_number_receipts(), ReportBranchData()._total_values_lottery_receipts('paid_number')),
        ("B44", BranchAsserts()._paid_amount_tickets(), ReportBranchData()._total_values_lottery_tickets('paid_amount')),
        ("F44", BranchAsserts()._paid_amount_receipts(), ReportBranchData()._total_values_lottery_receipts('paid_amount')),
        ("B45", BranchAsserts()._reward_tickets(), ReportBranchData()._total_values_lottery_tickets('reward')),
        ("F45", BranchAsserts()._reward_receipts(), ReportBranchData()._total_values_lottery_receipts('reward')),
        ("B47", BranchAsserts()._transfer_tickets(), ReportBranchData()._total_values_lottery_tickets('transfer')),
        ("F47", BranchAsserts()._transfer_receipts(), ReportBranchData()._total_values_lottery_receipts('transfer')),

        ("B49", BranchAsserts()._check_row('realization_tickets'), 'ДА'),
        ("G49", BranchAsserts()._check_row('realization_receipts'), 'ДА'),
        ("B51", ReportBranchData()._reward_of_two_tables(), BranchAsserts()._total_rewards()),
        ("B53", ReportBranchData()._transfer_of_two_tables(), BranchAsserts()._total_transfer()),

        # Отчет агента
        ("J41", AgentAsserts()._sold_number_tickets(), ReportAgentData()._total_values_lottery_tickets('sold_number')),
        (
        "N41", AgentAsserts()._sold_number_receipts(), ReportAgentData()._total_values_lottery_receipts('sold_number')),
        ("J42", AgentAsserts()._sold_amount_tickets(), ReportAgentData()._total_values_lottery_tickets('sold_amount')),
        (
        "N42", AgentAsserts()._sold_amount_receipts(), ReportAgentData()._total_values_lottery_receipts('sold_amount')),
        ("J43", AgentAsserts()._paid_number_tickets(), ReportAgentData()._total_values_lottery_tickets('paid_number')),
        (
        "N43", AgentAsserts()._paid_number_receipts(), ReportAgentData()._total_values_lottery_receipts('paid_number')),
        ("J44", AgentAsserts()._paid_amount_tickets(), ReportAgentData()._total_values_lottery_tickets('paid_amount')),
        (
        "N44", AgentAsserts()._paid_amount_receipts(), ReportAgentData()._total_values_lottery_receipts('paid_amount')),
        ("J45", AgentAsserts()._reward_tickets(), ReportAgentData()._total_values_lottery_tickets('reward')),
        ("N45", AgentAsserts()._reward_receipts(), ReportAgentData()._total_values_lottery_receipts('reward')),
        ("J47", AgentAsserts()._transfer_tickets(), ReportAgentData()._total_values_lottery_tickets('transfer')),
        ("N47", AgentAsserts()._transfer_receipts(), ReportAgentData()._total_values_lottery_receipts('transfer')),

        ("J49", AgentAsserts()._check_row('realization_tickets'), 'ДА'),
        ("O49", AgentAsserts()._check_row('realization_receipts'), 'ДА'),
        ("J51", ReportAgentData()._reward_of_two_tables(), AgentAsserts()._total_rewards()),
        ("J53", ReportAgentData()._transfer_of_two_tables(), AgentAsserts()._total_transfer())

    ]
    return array


