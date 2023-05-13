from main import CreatingFinalReport
import open_URL_in_one_click
from tkinter import messagebox
import tkinter

open_URL_info_text = "Для работы функционала, необходимо:\n" \
                     "1. Указать название файла без расширения, то есть без \".txt\"\n" \
                     "2. Файл должен находиться в папке \"Расхождения в отчетах\"(приоритет)\nили в одном месте с приложением"

create_report_info_text = "Для работы функционала, необходимо:\n" \
                          "1. Поместить приложение в одну папку с отчетами\n" \
                          "2. Изменить названия отчетов на:\n" \
                          "   а) Проданные билеты\n" \
                          "   б) Выплаченные выигрыши\n" \
                          "   в) Отчет для сверки\n" \
                          "   г) Отчет агента\n" \
                          "   д) Отчет агента для бестиражных лотерей\n" \
                          "   е) Отчет филиала\n" \
                          "   ж) Реестр выплаченных выигрышей"

# UI приложения
window = tkinter.Tk()
window.title("Trixie is glad to see you in her app! :-)")


def open_URL():
    """
    Открывает все URL ссылки, находящиеся в файле
    """
    file = file_entry.get()
    if len(file) >= 1:
        try:
            opening_URls.configure(
                open_URL_in_one_click.url_in_one_click(file)
            )
            tkinter.messagebox.showinfo(title="Info", message="Ссылки открыты. Хорошего дня!")
        except FileNotFoundError:
            tkinter.messagebox.showerror(title="Error", message=f"Файл \"{file}.txt\" не найден.")
    else:
        tkinter.messagebox.showwarning(title="Warning", message="Название файла не указано!")


def create_report():
    """
    Создает итоговый отчет
    """
    opening_URls.configure(
        CreatingFinalReport().calling_all_methods()
    )
    tkinter.messagebox.showinfo(title="Info", message="Итоговый отчет создан!")


# Блок внутри окна
frame = tkinter.Frame(window)
frame.pack()

"""Создание блока №1"""
opening_URls = tkinter.LabelFrame(frame, text="Открытие всех URL в файле")
opening_URls.grid(row=0, column=0, padx=20, pady=10)

# Информация по функционалу
opening_URls_label = tkinter.Label(opening_URls,
                                   text=open_URL_info_text,
                                   font=("Arial", 8, "italic"),  # font(шрифт, размер, курсив)
                                   justify="left")  # Выравнивание по левую сторону
opening_URls_label.grid(row=0, column=0, columnspan=2)

label_file_name = tkinter.Label(opening_URls, text="Укажите название файла:", font=("Arial", 8), justify="left")
label_file_name.grid(row=1, column=0, sticky="nw")

# Поле и кнопка для открытия всех ссылок в файле
file_entry = tkinter.Entry(opening_URls, font=("Arial", 7))
file_entry.grid(row=2, column=0, sticky="news")
file_entry.focus()
open_URL_button = tkinter.Button(opening_URls, text="Открытие всех URL в файле", command=open_URL)
open_URL_button.grid(row=2, column=1, sticky="news")

# Настройка всех детей внутри контейнера "user_info_frame"
for widget in opening_URls.winfo_children():
    widget.grid_configure(padx=10, pady=2)

"""Создание блока №2"""
creating_report = tkinter.LabelFrame(frame, text="Создание итогового отчета")
creating_report.grid(row=1, column=0, padx=20, pady=10, sticky="nw")

# Информация по функционалу
creating_report_label = tkinter.Label(creating_report, text=create_report_info_text, font=("Arial", 8, "italic"),
                                      justify="left")
creating_report_label.grid(row=0, column=0)
creating_report_label.grid_configure(padx=10, pady=2)

# TODO: Что это за бред :D
empty = tkinter.Label(creating_report, text="                                      ")
empty.grid(row=0, column=1)

create_report_button = tkinter.Button(creating_report, text="Создание итогового отчета", command=create_report)
create_report_button.grid(row=3, column=0, sticky="news", padx=10, pady=10, columnspan=2)
create_report_button.grid_configure(padx=10, pady=2)

window.mainloop()
