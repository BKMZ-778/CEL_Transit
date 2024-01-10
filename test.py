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
import pprint
import time

con = sl.connect('BAZA.db')

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
        df = pd.read_sql(f"Select party_numb, parcel_numb, parcel_plomb_numb, parcel_weight, decision_date from baza", con)
        print(df)
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


def update_data():
    con = sl.connect('BAZA.db')
    lis_parc_numb = pd.read_excel('list_upd.xlsx', header=None, engine='openpyxl')
    print(lis_parc_numb)
    with con:
        i = 0
        for parcel_numb in lis_parc_numb[0]:
            i += 1
            print(parcel_numb)
            con.execute(f"Update baza set pallet = 194 where parcel_plomb_numb = '{parcel_numb}'")
            print(i)


def test_API_insert_decision():
    parcel_list = ["RLV03175899",
                   "0120627328-0039-1",
                   "FS202337154711967414"]
    for parcel in parcel_list:
        body = {"registration_numb": "10716050/020923/П277237", "Event": "Выпуск тест",
                "parcel_numb": parcel,
                "Event": "Выпуск тест",
                "Event_comment": "тест",
                "Event_date": "2023-05-26 12:34:00",
                "Last_mile": "5Post for Cainiao"}
        headers = {'accept': 'application/json'}
        response = requests.post('http://164.132.182.145:5001/api/add_decision', json=body,  # http://164.132.182.145:5001
                                 headers={'accept': 'application/json'})
        print(response)

def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]

def chank_df():
    df = pd.read_excel('list_GBS.xlsx', header=None, engine='openpyxl')
    list_chanks = list(chunks(df[0], 25))
    print(list_chanks)
    for chank in list_chanks:
        list_of_dicts = []
        for i in chank:
            json_parc = {"HWBRefNumber": i}
            list_of_dicts.append(json_parc)
        print(list_of_dicts)
        print(len(list_of_dicts))


def GBS_request_events():
    df = pd.read_sql(f"Select parcel_numb from baza "
                     f"where party_numb LIKE '%GBS%' "  # where ID > (len(ID) - 200 000)
                     f"AND custom_status_short = 'ИЗЪЯТИЕ' "
                     f"AND custom_status != 'Return in process'", con).drop_duplicates(subset='parcel_numb')
    print(df)
    list_chanks = list(chunks(df['parcel_numb'], 25))
    #print(list_chanks)
    for chank in list_chanks:
        list_of_dicts = []
        for i in chank:
            json_parc = {"HWBRefNumber": i}
            list_of_dicts.append(json_parc)
        print(list_of_dicts)
        print(len(list_of_dicts))
        url = "https://api.gbs-broker.ru/"
        body = {
                 "jsonrpc": "2.0",
                 "method": "get_events_public",
                 "params": {
                     "Filter": {
                     "HWB": list_of_dicts

                     },
                     "TextLang": "ru"
                     },

                 "id": "c52f1b33-dg"
                }
        headers = {'Content-Type': 'application/json'}
        response = requests.post(url=url, json=body,  # http://164.132.182.145:5001
                                 headers=headers)

        print(response)
        json_events = response.json()
        print(json_events)
        result = json_events['result']
        """Event_triger = ['CI',
                        'AI',
                        'CR',
                        'CR2',
                        'CR3',
                        'PCD',
                        'ASF',
                        'HBA',
                        'HBA41',
                        'HBA42',
                        'HBA43',
                        'HBA44',
                        'HBA45',
                        'HBA46',
                        'NOID',
                        'DOCOK',
                        'RIC'
                        ]"""
        with con:
            for parcel_slot in result:
                parcel_numb = parcel_slot['HWBRefNumber']
                #event_code = parcel_slot['events'][0]['event_code']
                custom_status = parcel_slot['events'][0]['event_text']
                if 'clearance complete' in custom_status or 'Released by customs' in custom_status:
                    custom_status_short = 'ВЫПУСК'
                else:
                    custom_status_short = 'ИЗЪЯТИЕ'
                decision_date = parcel_slot['events'][0]['event_time']
                refuse_reason = parcel_slot['events'][0]['event_comment']
                con.execute(f"Update baza set "
                            f" custom_status = '{custom_status}',"
                            f" custom_status_short = '{custom_status_short}',"
                            f" decision_date = '{decision_date}',"
                            f" refuse_reason = '{refuse_reason}' where parcel_numb = '{parcel_numb}'")
                print('updated')

        """df = pd.json_normalize(result)
        print(df)
        writer = pd.ExcelWriter('GBS_json.xlsx', engine='xlsxwriter')
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()"""

def GBS_request_events_one():
    url = "https://api.gbs-broker.ru/"
    body = {
        "jsonrpc": "2.0",
        "method": "get_events_public",
        "params": {
            "Filter": {
                "HWB": [{'HWBRefNumber': 'AER002774133'}]

            },
            "TextLang": "ru"
        },

        "id": "c52f1b33-dg"
    }
    headers = {'Content-Type': 'application/json'}
    response = requests.post(url=url, json=body,  # http://164.132.182.145:5001
                             headers=headers)

    print(response)
    json_events = response.json()
    print(json_events)

def new():
    df_all_sent = pd.read_excel('C:/Users/User/Desktop/ДОКУМЕНТЫ/ОТГРУЖЕННОЕ/ALL_Manifests.xlsx')
    print(df_all_sent)


def Scarif_request_events():
    map_scarif_status = {'exds10': 'выпуск товаров без уплаты таможенных платежей',
                         'exds30': 'выпуск возвращаемых товаров разрешен',
                        'exds31': 'требуется уплата таможенных платежей',
                         'exds32': 'выпуск товаров разрешен, таможенные платежи уплачены',
                         'exds33': 'выпуск разрешён, ожидание по временному ввозу',
                      'exds40': 'разрешение на отзыв',
                      'exds70': 'продление срока выпуска',
                      'exds90': 'отказ в выпуске товаров',
                         'exds0': 'статус не определен'}
    con = sl.connect('BAZA.db')
    con.execute('pragma journal_mode=wal')
    df = pd.read_sql(f"Select parcel_numb from baza "
                     f"where party_numb NOT LIKE '%GBS%' "  # where ID > (len(ID) - 200 000)
                     f"AND custom_status_short = 'ИЗЪЯТИЕ'", con).drop_duplicates(subset='parcel_numb')
    for parcel_numb in df['parcel_numb']:
        url = f"https://cellog.deklarant.ru/api/external/parcel-status/{parcel_numb}"
        headers = {"api-token": "40e2f498-450c-4b9f-a509-7f4c8877a6ff"}
        try:
            response = requests.get(url=url,  # http://164.132.182.145:5001
                                     headers=headers)

            json_events = response.json()
            registration_numb = json_events['registryNumber']
            custom_status = json_events["externalStatus"]
            for key in map_scarif_status.keys():
                custom_status = custom_status.replace(key, map_scarif_status[key])
            print(custom_status)
            if 'Выпуск' in str(custom_status) or 'выпуск ' in str(custom_status):
                custom_status_short = 'ВЫПУСК'
            else:
                custom_status_short = 'ИЗЪЯТИЕ'
            decision_date = json_events["decisionAt"]
            refuse_reason = json_events["reasonMessage"]
            with con:

                con.execute(f"Update baza set"
                            f" registration_numb = '{registration_numb}',"
                            f" custom_status = '{custom_status}',"
                            f" custom_status_short = '{custom_status_short}',"
                            f" decision_date = '{decision_date}',"
                            f" refuse_reason = '{refuse_reason}'"
                            f"where parcel_numb = '{parcel_numb}'")
        except Exception as e:
            print(e)

Scarif_request_events()