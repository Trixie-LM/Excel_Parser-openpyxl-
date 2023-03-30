from sales_report import ReportSalesData, SalesAsserts, SaleListTicketsInArray
from payment_report import ReportPaymentsData, PaymentsAsserts, PaymentListTicketsInArray
from branch_report import ReportBranchData, BranchAsserts
from agent_report import ReportAgentData, AgentAsserts
from agent_report_about_noncirculated_tickets import ReportAgentNoncirculatedData, AgentNoncirculatedAsserts
import report_of_checking
import os

# Классы отчета продаж
report_sales_data = ReportSalesData()
sales_asserts = SalesAsserts()
sale_array_list = SaleListTicketsInArray()
# Классы отчета выплат
report_payments_data = ReportPaymentsData()
payments_asserts = PaymentsAsserts()
payment_array_list = PaymentListTicketsInArray()
# Классы отчета филиала
report_branch_data = BranchAsserts()
branch_asserts = ReportBranchData()
# Классы отчета агента
report_agent_data = ReportAgentData()
agent_asserts = AgentAsserts()
# Классы отчета агента о бестиражных билетах
report_agent_noncirculated_data = ReportAgentNoncirculatedData()
agent_noncirculated_asserts = AgentNoncirculatedAsserts()

name_folder = "Расхождения в отчетах"
if not os.path.isdir(name_folder):
     os.mkdir(name_folder)


def create_file(name, array):
    file = open(name_folder + '/' + name + ' [РАСХОЖДЕНИЯ].txt', 'w', encoding='utf-8')
    site = 'http://10.240.240.99/Lotteries_Trade11_Piganov/hs/Tickets/Status/Ticket/'

    for num in array:
        file.write(site + num + '\n')

    file.close()


def postconditions():
    # Для продаж
    if sales_asserts.counting_tickets('Электронный')[0] != report_of_checking.number_of_digital_sales():
        create_file('SALE. Электронные билеты', sale_array_list.digital_tickets())

    if sales_asserts.counting_tickets('Бумажный', 'Тиражная')[0] != report_of_checking.number_of_circulation_sales():
        create_file('SALE. Тиражные билеты', sale_array_list.draw_tickets())

    if sales_asserts.counting_tickets('Бумажный', 'Моментальная')[0] != report_of_checking.amount_of_instant_sales():
        create_file('SALE. Моментальные билеты', sale_array_list.instant_tickets())

    # Для выплат
    if payments_asserts.counting_tickets('Электронный')[0] != report_of_checking.number_of_digital_payments():
        create_file('PAYMENT. Электронные билеты', payment_array_list.digital_tickets())

    if payments_asserts.counting_tickets('Бумажный', 'Тиражная')[0] != report_of_checking.number_of_circulation_payments():
        create_file('PAYMENT. Тиражные билеты', payment_array_list.draw_tickets())

    if payments_asserts.counting_tickets('Бумажный', 'Моментальная')[0] != report_of_checking.amount_of_instant_payments():
        create_file('PAYMENT. Моментальные билеты', payment_array_list.instant_tickets())

postconditions()