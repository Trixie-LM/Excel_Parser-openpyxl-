from sales_report import SalesAsserts, SaleListTicketsInArray
from payment_report import PaymentsAsserts, PaymentListTicketsInArray
import report_of_checking
import os

# Классы отчета продаж
sales_asserts = SalesAsserts()
sale_array_list = SaleListTicketsInArray()
# Классы отчета выплат
payments_asserts = PaymentsAsserts()
payment_array_list = PaymentListTicketsInArray()

# Создание папки
name_folder = "Расхождения в отчетах"
if not os.path.isdir(name_folder):
    os.mkdir(name_folder)


def create_file(name, array):
    file = open('Расхождения в отчетах/' + name + '.txt', 'w', encoding='utf-8')
    site = 'http://10.240.240.99/Lotteries_Trade11_Piganov/hs/Tickets/Status/Ticket/'

    for num in array:
        file.write(site + num + '\n')

    file.close()


def postconditions():
    # Для продаж
    if sales_asserts.counting_tickets('Электронный')[0] != report_of_checking.number_of_digital_sales():
        create_file('_{SALE} Электронные билеты', sale_array_list.digital_tickets())

    if sales_asserts.counting_tickets('Бумажный', 'Тиражная')[0] != report_of_checking.number_of_circulation_sales():
        create_file('_{SALE} Тиражные билеты', sale_array_list.draw_tickets())

    if sales_asserts.counting_tickets('Бумажный', 'Моментальная')[0] != report_of_checking.amount_of_instant_sales():
        create_file('_{SALE} Моментальные билеты', sale_array_list.instant_tickets())

    # Для выплат
    if payments_asserts.counting_tickets('Электронный')[0] != report_of_checking.number_of_digital_payments():
        create_file('_{PAYMENT} Электронные билеты', payment_array_list.digital_tickets())

    if payments_asserts.counting_tickets('Бумажный', 'Тиражная')[
        0] != report_of_checking.number_of_circulation_payments():
        create_file('_{PAYMENT} Тиражные билеты', payment_array_list.draw_tickets())

    if payments_asserts.counting_tickets('Бумажный', 'Моментальная')[
        0] != report_of_checking.amount_of_instant_payments():
        create_file('_{PAYMENT} Моментальные билеты', payment_array_list.instant_tickets())


postconditions()
