# Импорт модулей
import pandas as pd
import os
import glob
import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar, DateEntry


def form_submit():
    global date1, date2
    date1 = dateFrom.get()
    date2 = dateTo.get()
    window.destroy()


# Создание диалогового окна
window = tk.Tk()
window.title('Выбор Даты')

# Расположение окна по центру экрана
window.update_idletasks()
s = window.geometry()
s = s.split('+')
s = s[0].split('x')
width_window = int(s[0])
height_window = int(s[1])
w = (window.winfo_screenwidth() // 2 - width_window)
h = (window.winfo_screenheight() // 2 - height_window)
window.geometry('+{}+{}'.format(w, h))

# Элементы окна
frame_add_form = tk.Frame(window, bg='black')
frame_add_form.grid(column=0, row=0, sticky='s')
textDateFrom = ttk.Label(frame_add_form, text='Дата с', width=25)
textDateTo = ttk.Label(frame_add_form, text='Дата по', width=25)
dateFrom = DateEntry(frame_add_form, width=22, foreground='black', normalforeground='black',
                   selectforeground='red', background='white',
                   date_pattern='dd-mm-YYYY')
dateTo = DateEntry(frame_add_form, width=22, foreground='black', normalforeground='black',
                   selectforeground='red', background='white',
                   date_pattern='dd-mm-YYYY')
btn_submit = ttk.Button(frame_add_form, text='Submit', command=form_submit)
# Расположение элементов
textDateFrom.grid(row=0, column=0, sticky='w', padx=25, pady=30)
textDateTo.grid(row=0, column=1, sticky='e', padx=25, pady=30)
dateFrom.grid(row=1, column=0, sticky='w', padx=25, pady=0)
dateTo.grid(row=1, column=1, sticky='e', padx=25, pady=0)
btn_submit.grid(row=2, column=0, columnspan=2, sticky='n', padx=25, pady=25)

window.mainloop()

# Сменим директорию
os.chdir('input')

# Список файлов Excel для объединения
xl_files = glob.glob('*.xlsx')

# Читаем каждую книгу объединяем в один датафрейм
combined = pd.concat([pd.read_excel(file)
                      for file in xl_files], ignore_index=True)

combined.to_excel('../output/combined.xlsx', index=False)

print(date1, date2)
