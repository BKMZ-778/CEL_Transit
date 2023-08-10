import time
import winsound
import pandas as pd
import numpy as np
from tkinter import ttk
import tkinter as tk
from tkinter import *
from tkinter import filedialog
import tkinter.messagebox as mb
import datetime
import openpyxl
from openpyxl.styles import Border, Side, PatternFill, Font, GradientFill, Alignment
from copy import copy
from openpyxl.utils.cell import range_boundaries
from openpyxl.worksheet.page import PageMargins
import random
import xlsxwriter
from tkinter import Button, Frame, Tk
import pandastable as pt
from pandastable import Table
import sqlite3 as sl

now = datetime.datetime.now().strftime("%d.%m.%Y")
plomb_list = []
refuse_parcels_list = []
parcels_list = []
options = {'align': 'w',
'cellbackgr': '#F4F4F3',
'cellwidth': 80,
'floatprecision': 2,
'thousandseparator': '',
'font': 'Arial',
'fontsize': 30,
'fontstyle': '',
'grid_color': '#ABB1AD',
'linewidth': 4,
'rowheight': 50,
'rowselectedcolor': '#E4DED4',
'textcolor': 'black'}
pd.set_option('display.max_columns', None)
dTDa1 = None

def load_decisions():
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name, sheet_name=0, engine='openpyxl')
    df['Вес брутто'] = df['Вес брутто'].replace(',', '.', regex=True).astype(float)
    df['Статус_ТО'] = df['Статус ТО']
    df_true_status_names = df.drop_duplicates()
    df['Статус_ТО'] = df['Статус_ТО'].replace(to_replace='Выпуск товаров без уплаты таможенных платежей', value='ВЫПУСК', regex=True)
    df['Статус_ТО'] = df['Статус_ТО'].replace(to_replace='Выпуск товаров разрешен, таможенные платежи уплачены', value='ВЫПУСК', regex=True)
    for cel in df['Статус_ТО']:
        if cel != 'ВЫПУСК':
            df.loc[df['Статус_ТО'] == cel, 'Статус_ТО'] = 'ИЗЪЯТИЕ'
    df['Пломба'] = df['Пломба'].astype(str)

    df = df.rename(columns={'Регистрационный номер реестра': 'registration_numb', 'Общая накладная': 'party_numb',
                            'Трек-номер': 'parcel_numb', 'Пломба': 'parcel_plomb_numb', 'Вес брутто': 'parcel_weight',
                            'Статус ТО': 'custom_status', 'Статус_ТО': 'custom_status_short', 'Дата решения': 'decision_date',
                            'Причина отказа в ТО': 'refuse_reason'})
    con = sl.connect('BAZA.db')
    # открываем базу
    with con:
        baza = con.execute("select count(*) from sqlite_master where type='table' and name='baza'")
        for row in baza:
            # если таких таблиц нет
            if row[0] == 0:
                # создаём таблицу
                with con:
                    con.execute("""
                                CREATE TABLE baza (
                                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                registration_numb VARCHAR(25) NOT NULL,
                                party_numb VARCHAR(20),
                                parcel_numb VARCHAR(20) NOT NULL UNIQUE ON CONFLICT REPLACE,
                                parcel_plomb_numb VARCHAR(20),
                                parcel_weight FLOAT,
                                custom_status VARCHAR(400),
                                custom_status_short VARCHAR(8),
                                decision_date VARCHAR(20),
                                refuse_reason VARCHAR(400),
                                pallet VARCHAR(10),
                                zone VARCHAR(5),
                                VH_status VARCHAR(10),
                                goods
                                );
                            """)
        df.to_sql('baza', con=con, if_exists='append', index=False)
        with con:
            data = con.execute("SELECT * FROM baza")
            for row in data:
                print(row)

def select_plomb(plomb):
    con = sl.connect('BAZA.db')
    with con:
        data = con.execute("SELECT * FROM baza WHERE parcel_plomb_numb = ?", (plomb,))
        for row in data:
            print(row)

def plomb_serch(plomb):
    global plomb_list
    global len_done_plomb
    global dTDa1
    if dTDa1:
        dTDa1.destroy()
    #plomb = str(entry_plomb.get())
    con = sl.connect('BAZA.db')
    df_plomb = pd.read_sql("SELECT * FROM baza WHERE parcel_plomb_numb = ?", con, params=(plomb,))
    #raws = df_plomb.loc[df_plomb['Статус ТО'] != 'ВЫПУСК', 'Статус ТО']
    raws1 = df_plomb[df_plomb["custom_status_short"] != 'ВЫПУСК'].index
    print(raws1)
    dTDa1 = tk.Toplevel(width='800', height=750)
    dTDa1.title('Накладные')
    dTDaPT = pt.Table(dTDa1, dataframe=df_plomb,
                      showtoolbar=True, showstatusbar=True,
                      width='800', height='750', maxcellwidth=500)
    dTDaPT.setRowHeight(60)
    dTDaPT.setFont()
    pt.config.apply_options(options, dTDaPT)
    #entry_plomb.delete(0, 'end')
    if len(list(raws1)) != 0:
        dTDaPT.setRowColors(rows=list(raws1), clr='red', cols='all')
        dTDaPT.show()
        print(df_plomb)
        print('!')
        winsound.PlaySound('Snd_Open_Bag.wav', winsound.SND_FILENAME)
        plomb_list.append(plomb)
        plomb_list = list(set(plomb_list))
        len_done_plomb = len(plomb_list)
        #label.config(text=f'обработанно: {len_done_plomb} \nиз {plomb_quont} (откзн:{refuse_plomb_quont})',
        #             fg='black', font=('Times', 20))
        print(len_done_plomb)
        #entry_plomb.grab_set()
        #entry_plomb.focus_set()
    elif df_plomb.empty:
        dTDa1.destroy()
        winsound.PlaySound('Snd_NoPlomb.wav', winsound.SND_FILENAME)
        #entry_plomb.focus_set()
    else:
        dTDaPT.show()
        print(df_plomb)
        winsound.PlaySound('Snd_All_Issue.wav', winsound.SND_FILENAME)
        print('ВЫПУСК')
        # window.grab_set()
        # window.focus_force()
        plomb_list.append(plomb)
        plomb_list = list(set(plomb_list))
        len_done_plomb = len(plomb_list)
        #label.config(text=f'обработанно: {len_done_plomb} \nиз {plomb_quont} (откзн:{refuse_plomb_quont})',
        #             fg='black', font=('Times', 20))
        #print(len_done_plomb)
        #entry_plomb.grab_set()
        #entry_plomb.focus_set()

load_decisions()