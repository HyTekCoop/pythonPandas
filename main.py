# Импорт модулей
import pandas as pd
import os
import glob
import tkinter as tk
from tkinter import ttk
from tkcalendar import Calendar, DateEntry
from pandas.io.excel import ExcelWriter


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
                     date_pattern='YYYY-mm-dd')
dateTo = DateEntry(frame_add_form, width=22, foreground='black', normalforeground='black',
                   selectforeground='red', background='white',
                   date_pattern='YYYY-mm-dd')
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
# combined['Дата продажи'] = pd.to_datetime(combined['Дата продажи'])  # возможно понадобиться, но не включать дата багуется и не отображается коректно в ексел
maskDate = (combined['Дата продажи'] >= date1) & (combined['Дата продажи'] <= date2)
combined = combined.loc[maskDate]
combined.to_excel('combinedDate.xlsx', index=False)


# Таблица продаж [sheet_name='продажи']

# Общая таблица по кол-ву продаж и возвратов
sale = combined['Обоснование для оплаты'] == 'Продажа'
refund = combined['Обоснование для оплаты'] == 'Возврат'
takeSale = combined.loc[sale]
saleCount = takeSale.iloc[:, [0, 5]].groupby('Артикул поставщика').count()
takeRefund = combined.loc[refund]
refundCount = takeRefund.iloc[:, [0, 5]].groupby('Артикул поставщика').count()
saleCount.rename(columns={'№': 'Кол-во продаж'}, inplace=True)
refundCount.rename(columns={'№': 'Кол-во возвратов'}, inplace=True)
countSaleTable = pd.concat([saleCount, refundCount], sort=False, axis=1)
countSaleTable = countSaleTable.fillna(0)
countSaleTable['Итого'] = countSaleTable['Кол-во продаж'] - countSaleTable['Кол-во возвратов']
summarySale = countSaleTable['Кол-во продаж'].sum()
sumaryRefund = countSaleTable['Кол-во возвратов'].sum()
summarySaleTable = countSaleTable['Итого'].sum()
summaryCountSaleTable = pd.Series(data={'Кол-во продаж': summarySale,
                                        'Кол-во возвратов': sumaryRefund,
                                        'Итого': summarySaleTable}, name='Общий итог')
countSaleTable = countSaleTable.append(summaryCountSaleTable, ignore_index=False)


# Таблица продаж
summarySalePrice = takeSale['Вайлдберриз реализовал Товар (Пр)'].sum()
summarySalePriceForSeller = takeSale['К перечислению Продавцу за реализованный Товар'].sum()
saleTable = pd.DataFrame({'Кол-во продаж': [summarySale],
                          'Вайлдберриз реализовал Товар (Пр)': [summarySalePrice],
                          'К перечислению Продавцу за реализованный Товар': [summarySalePriceForSeller],
                          'Комиссия ВБ': [summarySalePriceForSeller - summarySalePrice]})

# Таблица возвратов [sheet_name='возвраты']
summaryRefundPrice = takeRefund['Вайлдберриз реализовал Товар (Пр)'].sum()
summaryRefundPriceForSeller = takeRefund['К перечислению Продавцу за реализованный Товар'].sum()
# countRefundTable = countSaleTable[['Артикул поставщика', 'Кол-во возвратов']]
refundTable = pd.DataFrame({'Кол-во продаж': [sumaryRefund],
                          'Вайлдберриз реализовал Товар (Пр)': [summaryRefundPrice],
                          'К перечислению Продавцу за реализованный Товар': [summaryRefundPriceForSeller],
                          'Комиссия ВБ': [summaryRefundPriceForSeller - summaryRefundPrice]})
# takeRefundtable = takeRefund[['Вайлдберриз реализовал Товар (Пр)', 'К перечислению Продавцу за реализованный Товар']]

# Таблица сторно [sheet_name='сторно']
storno = combined['Обоснование для оплаты'] == 'Сторно продаж'
takeStorno = combined.loc[storno]
dataSummaryStorno = {'Кол-во': [takeStorno['Кол-во'].sum()],
                     'Вайлдберриз реализовал Товар (Пр)': [takeStorno['Вайлдберриз реализовал Товар (Пр)'].sum()],
                     'Цена розничная с учетом согласованной скидки': [takeStorno['Цена розничная с учетом согласованной скидки'].sum()],
                     'К перечислению Продавцу за реализованный Товар': [takeStorno['К перечислению Продавцу за реализованный Товар'].sum()]}
summaryTakeStorno = pd.DataFrame(dataSummaryStorno)

# Таблица брака [sheet_name='оплата брака']
defect = combined['Обоснование для оплаты'] == 'Оплата брака'
takeDefect = combined.loc[defect]
takeDefectTable = takeDefect[['Артикул поставщика', 'Кол-во', 'К перечислению Продавцу за реализованный Товар']]
summaryDefectTable = pd.Series(data={'Артикул поставщика': 'Общий итог',
                                   'Кол-во': takeDefectTable['Кол-во'].sum(),
                                   'К перечислению Продавцу за реализованный Товар': takeDefectTable['К перечислению Продавцу за реализованный Товар'].sum()})
takeDefectTable = takeDefectTable.append(summaryDefectTable, ignore_index=True)


# Таблица штрафы [sheet_name='штрафы']
fine = combined['Обоснование для оплаты'] == 'Штрафы'
takeFine = combined.loc[fine]
takeFineTable = takeFine[['Артикул поставщика', 'Количество возврата', 'Штрафы', 'Обоснование штрафов и доплат']]
summaryFineTable = pd.Series(data={'Артикул поставщика': 'Общий итог',
                                   'Количество возврата': takeFineTable['Количество возврата'].sum(),
                                   'Штрафы': takeFineTable['Штрафы'].sum()})
takeFineTable = takeFineTable.append(summaryFineTable, ignore_index=True)


# Запись в файл
with ExcelWriter('../output/combined1.xlsx', mode="a" if os.path.exists('../output/combined1.xlsx') else "w") as writer:
    countSaleTable.to_excel(writer, sheet_name='продажи', index=True)
    saleTable.to_excel(writer, sheet_name='продажи', index=False, startrow=len(countSaleTable.index) + 5)
    # поменять потом на refundCount в первой строке takeRefund
    refundCount.to_excel(writer, sheet_name='возвраты', index=True)
    refundTable.to_excel(writer, sheet_name='возвраты', index=False, startrow=len(refundCount.index) + 5)

    takeStorno.to_excel(writer, sheet_name='сторно', index=False)
    summaryTakeStorno.to_excel(writer, sheet_name='сторно', index=False, startrow=len(takeStorno.index) + 5)

    takeDefect.to_excel(writer, sheet_name='брак', index=False)
    takeDefectTable.to_excel(writer, sheet_name='брак', index=False, startrow=len(takeDefect.index) + 5)

    takeFine.to_excel(writer, sheet_name='штрафы', index=False)
    takeFineTable.to_excel(writer, sheet_name='штрафы', index=False, startrow=len(takeFine.index) + 5)
