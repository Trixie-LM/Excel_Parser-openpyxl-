from tkinter import messagebox
import tkinter
import os.path

total_report = 'C:/Users/Trixie_LM/Desktop/1C/Итоговый отчет.xlsx'
sales_report = 'C:/Users/Trixie_LM/Desktop/1C/Проданные билеты.xlsx'
payment_report = 'C:/Users/Trixie_LM/Desktop/1C/Выплаченные выигрыши.xlsx'
report_of_checking = 'C:/Users/Trixie_LM/Desktop/1C/Отчет для сверки.xlsm'
agent_report = 'C:/Users/Trixie_LM/Desktop/1C/Отчет агента.xlsx'
agent_noncirculated_report = 'C:/Users/Trixie_LM/Desktop/1C/Отчет агента для бестиражных лотерей.xlsx'
branch_report = 'C:/Users/Trixie_LM/Desktop/1C/Отчет филиала.xlsx'
payment_registry = 'C:/Users/Trixie_LM/Desktop/1C/Реестр выплаченных выигрышей.xlsx'

# total_report = './Итоговый отчет.xlsx'
# sales_report = './Проданные билеты.xlsx'
# payment_report = './Выплаченные выигрыши.xlsx'
# report_of_checking = './Отчет для сверки.xlsm'
# agent_report = './Отчет агента.xlsx'
# agent_noncirculated_report = './Отчет агента для бестиражных лотерей.xlsx'
# branch_report = './Отчет филиала.xlsx'
# payment_registry = './Реестр выплаченных выигрышей.xlsx'

list_reports = [sales_report, payment_report, report_of_checking,
                agent_report, agent_noncirculated_report, branch_report, payment_registry]

missing_reports = []

for report in list_reports:
    if not os.path.exists(report):
        missing_reports.append('* ' + report.replace('./', ''))

error_message = "\n".join(missing_reports)

if len(missing_reports) > 0:
    tkinter.messagebox.showerror(title="Error",
                                 message=f'Не найдены некоторые отчеты, а именно:\n\n{error_message}')
