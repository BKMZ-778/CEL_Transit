import pandas as pd
from tkinter import filedialog
import tkinter as tk
from tkinter import filedialog
import tkinter.messagebox as mb
import datetime
import numpy as np
import openpyxl
import xlsxwriter
import os
import sqlite3 as sl
import winsound
import requests


con = sl.connect('BAZA.db')
def load_decisions():
    df = pd.read_excel('Decisions.xlsx', sheet_name=0, engine='openpyxl')
    df['Вес брутто'] = df['Вес брутто'].replace(',', '.', regex=True).astype(float)
    df['Статус_ТО'] = df['Статус ТО']
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
                                refuse_reason VARCHAR(400)
                                );
                            """)

        df_isalready_in = pd.DataFrame()
        df_to_append = pd.DataFrame()
        for parcel_numb in df['parcel_numb']:
            row_isalready_in = pd.read_sql(f"Select * from baza where parcel_numb = '{parcel_numb}'", con)
            row = df.loc[df['parcel_numb'] == parcel_numb]
            if row_isalready_in.empty:
                #row.to_sql('baza', con=con, if_exists='append', index=False)
                df_to_append = df_to_append.append(row)
                print(row)
            else:
                custom_status_short = df.loc[df['parcel_numb'] == parcel_numb]['custom_status_short'].values[0]
                print(custom_status_short)
                custom_status = df.loc[df['parcel_numb'] == parcel_numb]['custom_status'].values[0]
                print(custom_status)
                con.execute(f"Update baza set custom_status = '{custom_status}',"
                            f" custom_status_short = '{custom_status_short}' where parcel_numb = '{parcel_numb}'")
            #df_isalready_in = df_isalready_in.append(row_isalready_in)
        print(df_to_append)
        df_to_append.to_sql('baza', con=con, if_exists='append', index=False)

def test(numb):
    con = sl.connect('BAZA.db')
    with con:
        df = pd.read_sql(f"Select * from baza where party_numb = '{numb}'", con)
        if df.empty:
            df = pd.read_sql(f"Select * from baza where pallet = '{numb}'", con)
            if df.empty:
                df = pd.read_sql(f"Select * from baza where parcel_plomb_numb = '{numb}'", con)
                if df.empty:
                    df = pd.read_sql(f"Select * from baza where parcel_numb = '{numb}'", con)

    print(df)
    writer = pd.ExcelWriter(f'df-{numb}.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
        writer.sheets['Sheet1'].set_column(0, 3, 10)
        writer.sheets['Sheet1'].set_column(1, 3, 20)
        writer.sheets['Sheet1'].set_column(2, 3, 20)
        writer.sheets['Sheet1'].set_column(3, 3, 20)
        writer.sheets['Sheet1'].set_column(4, 3, 30)
        writer.sheets['Sheet1'].set_column(5, 3, 20)

    writer.save()
def save_total_excel():
    con = sl.connect('BAZA.db')
    with con:
        df = pd.read_sql(f"Select * from baza", con)
    writer = pd.ExcelWriter('df.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    for column in df:
        column_width = max(df[column].astype(str).map(len).max(), len(column))
        col_idx = df.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
        writer.sheets['Sheet1'].set_column(0, 3, 10)
        writer.sheets['Sheet1'].set_column(1, 3, 20)
        writer.sheets['Sheet1'].set_column(2, 3, 20)
        writer.sheets['Sheet1'].set_column(3, 3, 20)
        writer.sheets['Sheet1'].set_column(4, 3, 30)
        writer.sheets['Sheet1'].set_column(5, 3, 20)

    writer.save()

def load_sample_manifest():
    file_name = filedialog.askopenfilename()
    df = pd.read_excel(file_name, sheet_name=0, engine='openpyxl')
    df = df[['Номер отправления ИМ', 'Номер пломбы', 'Наименование товара', '№ AWB']]
    df = df.rename(columns={'Номер отправления ИМ': 'parcel_numb',
                           'Номер пломбы': 'parcel_plomb_numb',
                           'Наименование товара': 'goods',
                           '№ AWB': 'party_numb'})
    print(df)
    df_group = df.groupby('parcel_numb', sort=False)['goods'].agg(','.join)
    print(df_group)
    df = pd.merge(df, df_group, how='left', left_on='parcel_numb', right_on='parcel_numb')
    df = df.drop_duplicates(subset=['parcel_numb'], keep='first')
    df['goods']= df['goods_y']
    df = df[['parcel_numb', 'parcel_plomb_numb', 'goods', 'party_numb']]
    df['custom_status'] = 'Unknown'
    df['custom_status_short'] = 'ИЗЪЯТИЕ'
    con = sl.connect('BAZA.db')
    # открываем базу
    writer = pd.ExcelWriter('df.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()


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
                                            registration_numb VARCHAR(25),
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
        df_to_append = pd.DataFrame()


        for parcel_numb in df['parcel_numb']:
            row_isalready_in = pd.read_sql(f"Select * from baza where parcel_numb = '{parcel_numb}'", con)
            row = df.loc[df['parcel_numb'] == parcel_numb]
            if row_isalready_in.empty:
                # row.to_sql('baza', con=con, if_exists='append', index=False)
                df_to_append = df_to_append.append(row)
                print(row)
            else:
                custom_status_short = df.loc[df['parcel_numb'] == parcel_numb]['custom_status_short'].values[0]
                print(custom_status_short)
                custom_status = df.loc[df['parcel_numb'] == parcel_numb]['custom_status'].values[0]
                print(custom_status)
                con.execute(f"Update baza set custom_status = '{custom_status}',"
                            f" custom_status_short = '{custom_status_short}' where parcel_numb = '{parcel_numb}'")
            # df_isalready_in = df_isalready_in.append(row_isalready_in)
        print(df_to_append)
        df_to_append.to_sql('baza', con=con, if_exists='append', index=False)
        winsound.PlaySound('resheniya_zagruzhenu.wav', winsound.SND_FILENAME)

def test_report_baza():
    con = sl.connect('BAZA-reports.db')
    with con:
        data = con.execute("""SELECT * FROM report""").fetchall()
        for row in data:
            print(row)

def test_API():
    con = sl.connect('BAZA.db')
    party_numb = '01053007-CEL-81'
    with con:
        df_request_decisions = pd.read_sql(f"SELECT parcel_numb FROM baza WHERE party_numb = '{party_numb}'", con)
    body = df_request_decisions.to_json(orient="records", indent=2)
    #body = {'parcel_numb': 'CEL7000012753CD'}
    print(body)
    headers = {'accept': 'application/json'}
    response = requests.post('http://164.132.182.145:5001/api/get_decisions',
                             json=body)  # http://127.0.0.1:5000  # 'http://164.132.182.145:5001/api/get_decisions'
    try:
        json_decisions = response.json()
        print(type(json_decisions))
        print(json_decisions)
        df_loaded_decisions = pd.DataFrame.from_records(json_decisions)
        print(df_loaded_decisions)
        with con:
            df_to_append = pd.DataFrame()
            for parcel_numb in df_loaded_decisions['parcel_numb']:
                row_isalready_in = pd.read_sql(f"Select * from baza where parcel_numb = '{parcel_numb}'", con)
                row = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb]
                if row_isalready_in.empty:
                    # row.to_sql('baza', con=con, if_exists='append', index=False)
                    df_to_append = df_to_append.append(row)
                    print('appended')
                else:
                    custom_status_short = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb]['custom_status_short'].values[0]
                    custom_status = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb]['custom_status'].values[0]
                    decision_date = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb]['decision_date'].values[0]
                    refuse_reason = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb]['refuse_reason'].values[0]

                    con.execute(f"Update baza set "
                                f" custom_status = '{custom_status}',"
                                f" custom_status_short = '{custom_status_short}',"
                                f" decision_date = '{decision_date}',"
                                f" refuse_reason = '{refuse_reason}' where parcel_numb = '{parcel_numb}'")
                    print('updated')
                #    row_isalready_in = pd.read_sql(f"Select * from baza where parcel_numb = '{parcel_numb}'", con)
                # df_isalready_in = df_isalready_in.append(row_isalready_in)
            df_to_append.to_sql('baza', con=con, if_exists='append', index=False)
            print(df_to_append)
    except:
        print(response)


test_API()