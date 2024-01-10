import pandas as pd
from tkinter import filedialog
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox as mb
import datetime
import numpy as np
import openpyxl
import xlsxwriter

def start():
    msg = "Выберите реестр"
    mb.showinfo("Проверка реестра", msg)
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name, sheet_name=0, header=None, engine='openpyxl',
                   skiprows=1, usecols='A:I, K:T, V', converters={7: str, 14: str, 15: str, 18: str})
    df.columns = ['Номер отправления ИМ', 'Фамилия', 'Имя', 'Отчество',
              'Адрес получателя', 'Город', 'Область', 'Индекс', 'Телефон',
              'Количество единиц товара', 'Наименование товара', 'Стоимость ед. товарной позиции',
              'Ссылка на товар', 'Серия паспорта', 'Номер паспорта', 'Дата выдачи', 'Дата рождения',
              'Идентификационный налоговый номер', 'Вес брутто (Вес позиции)', 'Клиент']

