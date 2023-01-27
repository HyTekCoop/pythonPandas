# Импорт модулей
import pandas as pd
import numpy as np
import os
import glob
import tkinter as tk

from tkinter import ttk, CENTER, Canvas
from pandas import DataFrame
from tkcalendar import DateEntry
from pandas.io.excel import ExcelWriter
from typing import Optional
from PIL import ImageTk

def info_window():
    windowInfo = tk.Tk()
    windowInfo.title('Инфо')
    windowInfo.geometry('+{}+{}'.format(w-305, h))

    canvas = Canvas (windowInfo, width=300, height=250, background='black')
    canvas.pack()
    #Стиль текста и сам текст
    canvas.create_text(150, 35, text='Контактная информация', fill='white', justify=CENTER, font="Verdana 15")
    canvas.create_text(150, 95, text='Номер телефона\n+79600035317', fill='white', justify=CENTER, font="Verdana 12")
    canvas.create_text(150, 145, text='Telegram\n@HyTekCoop', fill='white', justify=CENTER, font="Verdana 12")
    canvas.create_text(150, 235, text='© Черняев Александр', fill='white', justify=CENTER, font="Verdana 10")

def form_submit():
    executionText['text'] = 'Программа выполняется'
    executionText.update()
    global date1, date2
    date1 = dateFrom.get()
    date2 = dateTo.get()
    main()
    executionText['text'] = 'Программа выполнилась'

def main():
    # Сменим директорию
    # pathInput = 'input'
    if os.path.isdir('input'):
        os.chdir('input')

    # Список файлов Excel для объединения
    xl_files = glob.glob('*.xlsx')

    # Читаем каждую книгу объединяем в один датафрейм
    combined = pd.concat([pd.read_excel(file) for file in xl_files], ignore_index=True)
    # combined['Дата продажи'] = pd.to_datetime(combined['Дата продажи'])  # возможно понадобиться, но не включать дата багуется и не отображается коректно в ексел
    combinedLogistic = combined
    maskDate = (combined['Дата продажи'] >= date1) & (combined['Дата продажи'] <= date2)
    combined = combined.loc[maskDate]
    # combined.to_excel('../данные.xlsx')


    # Переменные для всех страниц кроме sheet_name='продажи'
    sale = combined['Обоснование для оплаты'] == 'Продажа'
    takeSale = combined.loc[sale]


    # Таблица продаж [sheet_name='продажи']
    sale1 = combined['Тип документа'] == 'Продажа'
    takeSale1 = combined.loc[sale1]
    takeSale1Table = pd.DataFrame({'Артикул поставщика': [],
                                   'Кол-во': [],
                                   'Вайлдберриз реализовал Товар (Пр)': [],
                                   'К перечислению Продавцу за реализованный Товар': [],
                                   'Цена розничная с учетом согласованной скидки': []})
    takeSale1Unique = takeSale1['Артикул поставщика'].unique()
    for i in takeSale1Unique:
        item = takeSale1.loc[takeSale1['Артикул поставщика'] == i]
        takeSale1Table.loc[len(takeSale1Table.index)] = [i, item['Кол-во'].sum(),
                                                         item['Вайлдберриз реализовал Товар (Пр)'].sum(),
                                                         item['К перечислению Продавцу за реализованный Товар'].sum(),
                                                         item['Цена розничная с учетом согласованной скидки'].sum()]
    summarySale = takeSale1Table['Кол-во'].sum()
    summarySalePrice = takeSale1Table['Вайлдберриз реализовал Товар (Пр)'].sum()
    summarySalePriceForSeller = takeSale1Table['К перечислению Продавцу за реализованный Товар'].sum()
    takeSale1Table.loc[len(takeSale1Table.index)] = ['Общий итог',
                                                     summarySale,
                                                     summarySalePrice,
                                                     summarySalePriceForSeller,
                                                     takeSale1Table['Цена розничная с учетом согласованной скидки'].sum()]
    takeSale1Table.insert(4,'Комиссия ВБ',takeSale1Table['Вайлдберриз реализовал Товар (Пр)'] - takeSale1Table['К перечислению Продавцу за реализованный Товар'], False)


    # Таблица Логистика продаж [sheet_name='логистика продаж']
    conditionLogisticSale = ((combinedLogistic['Обоснование для оплаты'] == 'Логистика') | (
                combinedLogistic['Обоснование для оплаты'] == 'Сторно продаж')) & (
                                        combinedLogistic['Количество возврата'] == 0)
    logisticSale = combinedLogistic.loc[conditionLogisticSale]
    newTakeSale = combined.loc[sale]

    newTakeSale['Srid'] = newTakeSale['Srid'].fillna(0)
    logisticSale['РезультатSrid'] = logisticSale['Srid'].isin(newTakeSale['Srid'])
    logisticSale['РезультатRid'] = logisticSale['Rid'].isin(newTakeSale['Rid'])
    newTakeSale['РезультатSrid'] = newTakeSale['Srid'].isin(logisticSale['Srid'])
    newTakeSale['РезультатRid'] = newTakeSale['Rid'].isin(logisticSale['Rid'])

    rowLogisticSaleForMonth = pd.DataFrame({'За выбранный месяц': []})
    logisticSaleForMonth = logisticSale.loc[logisticSale['Дата продажи'] >= date1]
    conditionFroMonth = (logisticSaleForMonth['РезультатSrid'] == True) | (logisticSaleForMonth['РезультатRid'] == True)
    logisticSaleForMonth = logisticSaleForMonth.loc[conditionFroMonth]
    logisticSaleForMonthTable = pd.DataFrame({'Артикул поставщика': [],
                                         'Кол-во': [],
                                         'Услуги по доставке товара покупателю': []})
    logisticSaleForMonthUnique = logisticSaleForMonth['Артикул поставщика'].unique()
    for i in logisticSaleForMonthUnique:
        item = logisticSaleForMonth.loc[logisticSaleForMonth['Артикул поставщика'] == i]
        item = item[['Артикул поставщика', 'Кол-во', 'Услуги по доставке товара покупателю']]
        logisticSaleForMonthTable.loc[len(logisticSaleForMonthTable.index)] = [i, item['Кол-во'].count(),
                                                                   item['Услуги по доставке товара покупателю'].sum()]
    logisticSaleForMonthTable.loc[len(logisticSaleForMonthTable.index)] = ['Общий итог', logisticSaleForMonthTable['Кол-во'].sum(),
                                                               logisticSaleForMonthTable['Услуги по доставке товара покупателю'].sum()]

    rowLogisticSaleForTwoMonth = pd.DataFrame({'За 2 предыдущих месяца': []})
    logisticSaleForTwoMonth = logisticSale.loc[logisticSale['Дата продажи'] < date1]
    conditionFroTwoMonth = (logisticSaleForTwoMonth['РезультатSrid'] == True) | (
                logisticSaleForTwoMonth['РезультатRid'] == True)
    logisticSaleForTwoMonth = logisticSaleForTwoMonth.loc[conditionFroTwoMonth]
    logisticSaleForTwoMonthTable = pd.DataFrame({'Артикул поставщика': [],
                                         'Кол-во': [],
                                         'Услуги по доставке товара покупателю': []})
    logisticSaleForTwoMonthUnique = logisticSaleForTwoMonth['Артикул поставщика'].unique()
    for i in logisticSaleForTwoMonthUnique:
        item = logisticSaleForTwoMonth.loc[logisticSaleForTwoMonth['Артикул поставщика'] == i]
        item = item[['Артикул поставщика', 'Кол-во', 'Услуги по доставке товара покупателю']]
        logisticSaleForTwoMonthTable.loc[len(logisticSaleForTwoMonthTable.index)] = [i, item['Кол-во'].count(),
                                                                   item['Услуги по доставке товара покупателю'].sum()]
    logisticSaleForTwoMonthTable.loc[len(logisticSaleForTwoMonthTable.index)] = ['Общий итог', logisticSaleForTwoMonthTable['Кол-во'].sum(),
                                                               logisticSaleForTwoMonthTable['Услуги по доставке товара покупателю'].sum()]

    rowSaleNotFound = pd.DataFrame({'Ненайденная логистика': []})
    conditionForSaleNotFound = ((newTakeSale['РезультатSrid'] == False) & (newTakeSale['РезультатRid'] == False))
    logisticSaleNotFound = newTakeSale.loc[conditionForSaleNotFound]
    logisticSaleNotFoundTable = pd.DataFrame({'Артикул поставщика': [],
                                              'Кол-во': []})
    logisticSaleNotFoundUnique = logisticSaleNotFound['Артикул поставщика'].unique()
    for i in logisticSaleNotFoundUnique:
        item = logisticSaleNotFound.loc[logisticSaleNotFound['Артикул поставщика'] == i]
        item = item[['Артикул поставщика', 'Кол-во']]
        logisticSaleNotFoundTable.loc[len(logisticSaleNotFoundTable.index)] = [i, item['Кол-во'].count()]
    logisticSaleNotFoundTable.loc[len(logisticSaleNotFoundTable.index)] = ['Общий итог',
                                                                           logisticSaleNotFoundTable['Кол-во'].sum()]


    # Таблица возвратов [sheet_name='возвраты']
    # refund = combined['Тип документа'] == 'Возврат'
    refund = (combined['Обоснование для оплаты'] == 'Возврат') | (combined['Обоснование для оплаты'] == 'Корректный возврат')
    takeRefund = combined.loc[refund]
    takeRefundTable = pd.DataFrame({'Артикул поставщика': [],
                                   'Кол-во': [],
                                   'Вайлдберриз реализовал Товар (Пр)': [],
                                   'К перечислению Продавцу за реализованный Товар': [],
                                   'Цена розничная с учетом согласованной скидки': []})
    takeRefundUnique = takeRefund['Артикул поставщика'].unique()
    for i in takeRefundUnique:
        item = takeRefund.loc[takeRefund['Артикул поставщика'] == i]
        takeRefundTable.loc[len(takeRefundTable.index)] = [i, item['Кол-во'].sum(),
                                                         item['Вайлдберриз реализовал Товар (Пр)'].sum(),
                                                         item['К перечислению Продавцу за реализованный Товар'].sum(),
                                                         item['Цена розничная с учетом согласованной скидки'].sum()]
    summaryRefund = takeRefundTable['Кол-во'].sum()
    summaryRefundPrice = takeRefundTable['Вайлдберриз реализовал Товар (Пр)'].sum()
    summaryRefundPriceForSeller = takeRefundTable['К перечислению Продавцу за реализованный Товар'].sum()
    takeRefundTable.loc[len(takeRefundTable.index)] = ['Общий итог',
                                                     summaryRefund,
                                                     summaryRefundPrice,
                                                     summaryRefundPriceForSeller,
                                                     takeRefundTable['Цена розничная с учетом согласованной скидки'].sum()]
    takeRefundTable.insert(4,'Комиссия ВБ',takeRefundTable['Вайлдберриз реализовал Товар (Пр)'] - takeRefundTable['К перечислению Продавцу за реализованный Товар'], False)


    # Таблица Логистика возвратов [sheet_name='логистика возвратов']
    logisticRefund = (combined['Обоснование для оплаты'] == 'Логистика') & (combined['Количество возврата'] > 0)
    takeLogisticRefund = combined.loc[logisticRefund]
    logisticRefundTable = pd.DataFrame({'Артикул поставщика': [],
                                         'Кол-во': [],
                                         'Услуги по доставке товара покупателю': []})

    logisticRefundUnique = takeLogisticRefund['Артикул поставщика'].unique()
    for i in logisticRefundUnique:
        item = takeLogisticRefund.loc[takeLogisticRefund['Артикул поставщика'] == i]
        item = item[['Артикул поставщика', 'Кол-во', 'Услуги по доставке товара покупателю']]
        logisticRefundTable.loc[len(logisticRefundTable.index)] = [i, item['Кол-во'].count(),
                                                                   item['Услуги по доставке товара покупателю'].sum()]

    summaryLogisticRefund = takeLogisticRefund['Количество возврата'].sum()
    summaryLogisticRefundForSeller = takeLogisticRefund['Услуги по доставке товара покупателю'].sum()
    logisticRefundTable.loc[len(logisticRefundTable.index)] = ['Общий итог', logisticRefundTable['Кол-во'].sum(),
                                                               logisticRefundTable['Услуги по доставке товара покупателю'].sum()]
    # Таблица сторно продаж [sheet_name='сторно продаж']
    stornoSale = combined['Обоснование для оплаты'] == 'Сторно продаж'
    takeStornoSale = combined.loc[stornoSale]
    stornoSaleTable = pd.DataFrame({'Артикул поставщика': [],
                                'Кол-во': [],
                                'Вайлдберриз реализовал Товар (Пр)': [],
                                'К перечислению Продавцу за реализованный Товар': [],
                                'Цена розничная с учетом согласованной скидки': []})

    takeStornoSaleUnique = takeStornoSale['Артикул поставщика'].unique()
    for i in takeStornoSaleUnique:
        item = takeStornoSale.loc[takeStornoSale['Артикул поставщика'] == i]
        stornoSaleTable.loc[len(stornoSaleTable.index)] = [i, item['Кол-во'].sum(),
                                                   item['Вайлдберриз реализовал Товар (Пр)'].sum(),
                                                   item['К перечислению Продавцу за реализованный Товар'].sum(),
                                                   item['Цена розничная с учетом согласованной скидки'].sum()]
    summaryStornoSale = stornoSaleTable['Кол-во'].sum()
    summaryStornoSalePrice = stornoSaleTable['Вайлдберриз реализовал Товар (Пр)'].sum()
    summaryStornoSalePriceForSeller = stornoSaleTable['К перечислению Продавцу за реализованный Товар'].sum()
    stornoSaleTable.loc[len(stornoSaleTable.index)] = ['Общий итог',
                                                       summaryStornoSale,
                                                       summaryStornoSalePrice,
                                                       summaryStornoSalePriceForSeller,
                                                       stornoSaleTable['Цена розничная с учетом согласованной скидки'].sum()]
    stornoSaleTable.insert(4, 'Комиссия ВБ',
                           stornoSaleTable['Вайлдберриз реализовал Товар (Пр)'] - stornoSaleTable['К перечислению Продавцу за реализованный Товар'], False)

    # Таблица сторно возвратов [sheet_name='сторно возратов']
    stornoRefund = combined['Обоснование для оплаты'] == 'Сторно возвратов'
    takeStornoRefund = combined.loc[stornoRefund]
    stornoRefundTable = pd.DataFrame({'Артикул поставщика': [],
                                    'Кол-во': [],
                                    'Вайлдберриз реализовал Товар (Пр)': [],
                                    'К перечислению Продавцу за реализованный Товар': [],
                                    'Цена розничная с учетом согласованной скидки': []})

    takeStornoRefundUnique = takeStornoRefund['Артикул поставщика'].unique()
    for i in takeStornoRefundUnique:
        item = takeStornoRefund.loc[takeStornoRefund['Артикул поставщика'] == i]
        stornoRefundTable.loc[len(stornoRefundTable.index)] = [i, item['Кол-во'].sum(),
                                                           item['Вайлдберриз реализовал Товар (Пр)'].sum(),
                                                           item['К перечислению Продавцу за реализованный Товар'].sum(),
                                                           item['Цена розничная с учетом согласованной скидки'].sum()]
    summaryStornoRefund = stornoRefundTable['Кол-во'].sum()
    summaryStornoRefundPrice = stornoRefundTable['Вайлдберриз реализовал Товар (Пр)'].sum()
    summaryStornoRefundPriceForSeller = stornoRefundTable['К перечислению Продавцу за реализованный Товар'].sum()
    stornoRefundTable.loc[len(stornoRefundTable.index)] = ['Общий итог',
                                                       summaryStornoRefund,
                                                       summaryStornoRefundPrice,
                                                       summaryStornoRefundPriceForSeller,
                                                        stornoRefundTable['Цена розничная с учетом согласованной скидки'].sum()]
    stornoRefundTable.insert(4, 'Комиссия ВБ',
                           stornoRefundTable['Вайлдберриз реализовал Товар (Пр)'] - stornoRefundTable[
                               'К перечислению Продавцу за реализованный Товар'], False)

    # Таблица сторно [sheet_name='Оплата потерянного товара']
    lostItem = combined['Обоснование для оплаты'] == 'Оплата потерянного товара'
    takeLostItem = combined.loc[lostItem]
    lostItemTable = pd.DataFrame({'Артикул поставщика': [],
                                  'Кол-во': [],
                                  'Вайлдберриз реализовал Товар (Пр)': [],
                                  'К перечислению Продавцу за реализованный Товар': [],
                                  'Цена розничная с учетом согласованной скидки': []})

    takeLostItemUnique = takeLostItem['Артикул поставщика'].unique()
    for i in takeLostItemUnique:
        item = takeLostItem.loc[takeLostItem['Артикул поставщика'] == i]
        lostItemTable.loc[len(lostItemTable.index)] = [i, item['Кол-во'].sum(),
                                                               item['Вайлдберриз реализовал Товар (Пр)'].sum(),
                                                               item[
                                                                   'К перечислению Продавцу за реализованный Товар'].sum(),
                                                               item[
                                                                   'Цена розничная с учетом согласованной скидки'].sum()]
    summaryLostItem = lostItemTable['Кол-во'].sum()
    summaryLostItemPrice = lostItemTable['Вайлдберриз реализовал Товар (Пр)'].sum()
    summaryLostItemPriceForSeller = lostItemTable['К перечислению Продавцу за реализованный Товар'].sum()
    lostItemTable.loc[len(lostItemTable.index)] = ['Общий итог',
                                                           summaryLostItem,
                                                           summaryLostItemPrice,
                                                           summaryLostItemPriceForSeller,
                                                           lostItemTable[
                                                               'Цена розничная с учетом согласованной скидки'].sum()]
    lostItemTable.insert(4, 'Комиссия ВБ',
                             lostItemTable['Вайлдберриз реализовал Товар (Пр)'] - lostItemTable[
                                 'К перечислению Продавцу за реализованный Товар'], False)


    # Таблица поставок [sheet_name='поставки']
    suppliesTable: list[Optional[DataFrame]] = []
    suppliesUnique = takeSale['Номер поставки'].unique()
    for supplies in suppliesUnique:
        currentSupplie = takeSale.loc[takeSale['Номер поставки'] == supplies]
        currentSupplie = currentSupplie[['Артикул поставщика', 'Кол-во']]
        currentSupplie = currentSupplie.groupby('Артикул поставщика').count()
        currentSupplie.reset_index(inplace=True)
        currentSupplie.rename(columns={'Артикул поставщика': 'Номер поставки', 'Кол-во': supplies}, inplace=True)
        suppliesTable.append(pd.concat([pd.DataFrame({'Номер поставки': ['Артикул поставщика'],
                                                      supplies: ['Кол-во']}), currentSupplie], ignore_index=False,
                                       axis=0))

    # Таблица брака [sheet_name='брака']
    defect = combined['Обоснование для оплаты'] == 'Оплата брака'
    takeDefect = combined.loc[defect]
    takeDefectTable = takeDefect[['Артикул поставщика', 'Кол-во', 'Вайлдберриз реализовал Товар (Пр)']]
    summaryDefect = takeDefectTable['Кол-во'].sum()
    summaryDefectPrice = takeDefectTable['Вайлдберриз реализовал Товар (Пр)'].sum()
    summaryDefectTable = pd.Series(data={'Артикул поставщика': 'Общий итог',
                                         'Кол-во': summaryDefect,
                                         'Вайлдберриз реализовал Товар (Пр)': summaryDefectPrice})
    takeDefectTable = takeDefectTable.append(summaryDefectTable, ignore_index=True)

    # Таблица штрафы [sheet_name='штрафы']
    fine = combined['Обоснование для оплаты'] == 'Штрафы'
    takeFine = combined.loc[fine]

    if ('Общая сумма штрафов' in takeFine.columns) & ('Штрафы' in takeFine):
        takeFine['Общая сумма штрафов'] = takeFine['Общая сумма штрафов'].fillna(0)
        takeFine['Штрафы'] = takeFine['Штрафы'].fillna(0)
        takeFine.rename(columns = {'Штрафы' : 'Штрафы1'}, inplace=True)
        takeFine.insert(27, 'Штрафы', takeFine['Общая сумма штрафов'] + takeFine['Штрафы1'], False)
    elif 'Общая сумма штрафов' in takeFine.columns:
        takeFine.rename(columns={'Общая сумма штрафов': 'Штрафы'}, inplace=True)

    takeFineTable = takeFine[['Артикул поставщика', 'Количество возврата', 'Штрафы', 'Обоснование штрафов и доплат']]
    summaryFineRuturnTable = takeFineTable['Количество возврата'].sum()
    summaryFinePrice = takeFineTable['Штрафы'].sum()
    summaryFineTable = pd.Series(data={'Артикул поставщика': 'Общий итог',
                                       'Количество возврата': summaryFineRuturnTable,
                                       'Штрафы': summaryFinePrice})
    takeFineTable = takeFineTable.append(summaryFineTable, ignore_index=True)


    # Таблица Итоговая [sheet_name='ОПУ']
    summaryOpyQuantity = summarySale - summaryRefund - summaryStornoSale - summaryStornoRefund - summaryDefect - summaryLostItem
    summaryOpyPrice = summarySalePrice - summaryRefundPriceForSeller - \
                      summaryStornoSalePriceForSeller - summaryStornoRefundPriceForSeller - \
                      summaryDefectPrice - summaryLostItemPrice
    opy = pd.DataFrame({'Наименование строки ОПУ': [np.nan, 'Выручка', 'Возвраты', 'Сторно продаж', 'Сторно возвратов',
                                                    'Брак', 'Потерянный товар', 'Итого выручка', np.nan,
                                                    'Себестоимость', 'доставка до мск', 'фулфилмент', 'логистика  ВБ',
                                                    'комиссия ВБ', 'стоимость хранения', 'стоимость платной приемки',
                                                    'логистика возвратов', 'прочие удержания', 'штрафы',
                                                    'Валовая прибыль'],
                        'Пояснение': [np.nan, '*комиссия выделена отдельно', '*с учетом комиссии вб',
                                      '*с учетом комиссии вб', np.nan, np.nan, np.nan, np.nan, np.nan, np.nan,
                                      '*берем из файла себестоимость', '*берем из файла себестоимость',
                                      np.nan, '*выделение комиссии по продажам', '*возможно взять только понедельно',
                                      '*возможно взять только понедельно', np.nan, '*возможно взять только понедельно',
                                      np.nan, np.nan],
                        'Кол-во': [np.nan, summarySale, summaryRefund, summaryStornoSale, summaryStornoRefund,
                                   summaryDefect, summaryLostItem, summaryOpyQuantity, np.nan,
                                   np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, np.nan, summaryLogisticRefund,
                                   np.nan, summaryFineRuturnTable, np.nan],
                        'В руб.': [np.nan, summarySalePrice, summaryRefundPriceForSeller, summaryStornoSalePriceForSeller,
                                   summaryStornoRefundPriceForSeller, summaryDefectPrice, summaryLostItemPrice,
                                   summaryOpyPrice, np.nan, np.nan, np.nan, np.nan, np.nan,
                                   summarySalePrice - summarySalePriceForSeller, np.nan, np.nan,
                                   summaryLogisticRefundForSeller, np.nan, summaryFinePrice, np.nan]})

    # Удаление из папки существующего файла Отчеты
    pathReport = '../output/Отчет.xlsx'
    if os.path.isfile(pathReport):
        os.remove(pathReport)

    # Запись в файл
    # Если включить закомментированные строки будут записываться еще и данные по которым сделан рассчет
    with ExcelWriter(pathReport, mode="a" if os.path.exists(pathReport) else "w") as writer:
        takeSale1Table.to_excel(writer, sheet_name='продажи', index=False)
        # takeSale.to_excel(writer, sheet_name='продажи', index=False, startrow=len(takeSale1Table.index) + 5)

        rowSaleNotFound.to_excel(writer, sheet_name='логистика продаж', index=False, startrow=0)
        logisticSaleNotFoundTable.to_excel(writer, sheet_name='логистика продаж', index=False, startrow=1)
        length = len(logisticSaleNotFoundTable.index)
        rowLogisticSaleForMonth.to_excel(writer, sheet_name='логистика продаж', index=False, startrow=length + 7)
        logisticSaleForMonthTable.to_excel(writer, sheet_name='логистика продаж', index=False, startrow=length + 8)
        rowLogisticSaleForTwoMonth.to_excel(writer, sheet_name='логистика продаж', index=False, startrow=length + 7, startcol=8)
        logisticSaleForTwoMonthTable.to_excel(writer, sheet_name='логистика продаж', index=False, startrow=length + 8, startcol=8)
        # logisticSaleForTwoMonth.to_excel(writer, sheet_name='логистика продаж', index=False, startrow=len(logisticSaleForTwoMonthTable.index) + 5, startcol=8)

        takeRefundTable.to_excel(writer, sheet_name='возвраты', index=False)
        # takeRefund.to_excel(writer, sheet_name='возвраты', index=False, startrow=len(takeRefundTable.index) + 5)

        logisticRefundTable.to_excel(writer, sheet_name='логистика возвратов', index=False)

        stornoSaleTable.to_excel(writer, sheet_name='сторно продаж', index=False)
        # takeStornoSale.to_excel(writer, sheet_name='сторно продаж', index=False, startrow=len(stornoSaleTable.index) + 5)

        stornoRefundTable.to_excel(writer, sheet_name='сторно возвратов', index=False)
        # takeStornoRefund.to_excel(writer, sheet_name='сторно возвратов', index=False, startrow=len(stornoRefundTable.index) + 5)

        lostItemTable.to_excel(writer, sheet_name='оплата потерянного товара', index=False)
        # takeLostItem.to_excel(writer, sheet_name='оплата потерянного товара', index=False, startrow=len(lostItemTable.index) + 5)


        row = 0
        for i in range(len(suppliesTable)):
            suppliesTable[i].to_excel(writer, sheet_name='поставки', index=False, startrow=row)
            row = row + len(suppliesTable[i].index) + 2

        takeDefectTable.to_excel(writer, sheet_name='брак', index=False)
        # takeDefect.to_excel(writer, sheet_name='брак', index=False, startrow=len(takeDefect.index) + 5)

        takeFineTable.to_excel(writer, sheet_name='штрафы', index=False)
        # takeFine.to_excel(writer, sheet_name='штрафы', index=False, startrow=len(takeFine.index) + 5)

        opy.to_excel(writer, sheet_name='ОПУ', index=False)


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
executionText = ttk.Label(frame_add_form, text='', width=25, justify=CENTER, font="Verdana 12",
                           background='BLACK', foreground='white', padding=(0,0,-30,0))
btn_submit = ttk.Button(frame_add_form, text='Рассчитать', command=form_submit)
infoIcon = ImageTk.PhotoImage(file="image/icons8-info-squared-35.png")
btn_info = ttk.Button(frame_add_form, image=infoIcon, command=info_window, width=30, padding=-10)

# Расположение элементов
textDateFrom.grid(row=0, column=0, sticky='w', padx=25, pady=30)
textDateTo.grid(row=0, column=1, sticky='e', padx=25, pady=30)
dateFrom.grid(row=1, column=0, sticky='w', padx=25, pady=0)
dateTo.grid(row=1, column=1, sticky='e', padx=25, pady=0)
executionText.grid(row=2, column=0, columnspan=2, sticky='n', padx=25, pady=(25,0))
btn_submit.grid(row=3, column=0, columnspan=2, sticky='n', padx=25, pady=25)
btn_info.grid(row=0, column=0, sticky='nw', padx=0, pady=0)
window.mainloop()
