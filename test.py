import json

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
import base64
from base64 import b64encode
import hashlib
import os



pd.set_option('display.max_columns', None)
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
        df = pd.read_sql(f"Select * from baza where ID > 2539878", con)
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
                "HWB": [{'HWBRefNumber': '99880012029454'}]

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
    filename = filedialog.askopenfilename()
    map_scarif_status = {'exds10': 'выпуск товаров без уплаты таможенных платежей',
                         'exds30': 'выпуск возвращаемых товаров разрешен',
                        'exds31': 'требуется уплата таможенных платежей',
                         'exds32': 'выпуск товаров разрешен, таможенные платежи уплачены',
                         'exds33': 'выпуск разрешён, ожидание по временному ввозу',
                      'exds40': 'разрешение на отзыв',
                      'exds70': 'продление срока выпуска',
                      'exds90': 'отказ в выпуске товаров',
                         'exds0': 'статус не определен'}
    df = pd.read_excel(filename, sheet_name=0, engine='openpyxl', usecols='A, AO', header=None)
    for parcel_numb in df[0]:
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

def send_parcel_status():
    data = {"PostingNumber": "CEL6000269115CD", "TrackingNumber": "CEL6000269115CD",
            "Data": [{"track_code": "DS","datetime": "2024/1/12 16:20:20","location": "Россия", "description": "Выпуск товаров без уплаты таможенных платежей"}]}
    data_str = str(data).replace("'", '"').replace(", ", ",")

    m = hashlib.md5()
    m.update(data_str.encode('utf-8'))
    result = base64.urlsafe_b64encode(m.hexdigest().encode('utf-8')).decode('utf-8') #b64encode(m.hexdigest().encode('utf-8'))
    print(result)
    url = ("http://hccd.rtb56.com/webservice/edi/TrackService.ashx?code=ADDCUSTOMSCLEARANCETRACK"
           + f'&data={data_str}' + f'&sign={str(result)}')

    print(url)
    url_test = 'http://hccd.rtb56.com/webservice/edi/TrackService.ashx?code=ADDCUSTOMSCLEARANCETRACK&data={"PostingNumber": "33404523118556","TrackingNumber": "FG240106000137","Data": [{"track_code": "DS","datetime": "2024/1/3 16:20:20","location": null,"description": "Отпусти."},{"track_code": "DS","datetime": "2024/1/3 16:20:20","location": null,"description": "Не выпущено:Предоставление паспорта"}]}&sign=MzlkMTdhZmIxYTlmNWE1YmIyNmI3ZmM4ZTlmOTVkNzA='
    print(url_test)
    response = requests.get(url)
    print(response.text)


def hesh_data():
    data = {'PostingNumber': '', 'TrackingNumber': 'CEL6000269113CD',
            'Data': [{'track_code': '', 'datetime': '03.01.2024 16:20:20', 'location': 'Ussuriysk', 'description': 'Выпуск товаров без уплаты таможенных платежей'}]}
    data_md5 = hashlib.md5(str(data).encode('utf-8')).hexdigest()
    data_base64 = b64encode(data_md5.encode())
    print(data_base64)

def encode_sing():
    sing = 'N2Q2MDBhYTdmNzVhZTEwOWNiYTJlZDM0NGMyNTU0ODM='

def all_issyes_to_exl():
    con = sl.connect('BAZA.db')
    con_vect = sl.connect('VECTORS.db')
    with con:
        df_parties = pd.read_sql("SELECT * FROM baza WHERE party_numb in ('9542101-OZON-220')", con)
        print(df_parties)
    with con_vect:
        df_vectors = pd.read_sql("SELECT * FROM vectors WHERE party_numb in ('9542101-OZON-220')", con_vect)

    #df_vectors.drop_duplicates(subset='parcel_plomb_numb')
    print(df_vectors)
    df_merged = pd.merge(df_parties, df_vectors, how='left', left_on='parcel_numb', right_on='parcel_numb')
    print(df_merged)
    writer = pd.ExcelWriter(f'www1.xlsx', engine='xlsxwriter')
    df_merged.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()


def Scarif_request_few_parcel():
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
    List_of_parcel = ["FS20236U154711477994",
"CEL7000020424CD",
"CEL7000026724CD",
"CEL7000032957CD"]
    dict_result = {}
    for parcel_numb in List_of_parcel:
        try:
            print(parcel_numb)
            url = f"https://cellog.deklarant.ru/api/external/parcel-status/{parcel_numb}"
            headers = {"api-token": "40e2f498-450c-4b9f-a509-7f4c8877a6ff"}

            response = requests.get(url=url,  # http://164.132.182.145:5001
                                     headers=headers)

            print(response)
            json_events = response.json()
            print(json_events)
            registration_numb = json_events['registryNumber']
            print(registration_numb)
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

            dict_result["parcel_numb"] = parcel_numb
            dict_result["reg_numb"] = registration_numb

        except Exception as e:
            print(e)

    print(dict_result)

def api_track718(track):
    cel_api_key = "e0fca820-c3dc-11ee-b960-bdfb353c94dc"
    url = "https://apigetway.track718.net/v2/tracking/query"
    headers = {"Content-Type": "application/json",
    "Track718-API-Key": f"{cel_api_key}"}

    params = [{"trackNum": track, "code": "gps-truck"}]
    respons = requests.post(url=url, headers=headers, json=params)


    print(respons.status_code)
    print(respons)
    print(respons.json())

def api_track718_add_track(track):
    cel_api_key = "e0fca820-c3dc-11ee-b960-bdfb353c94dc"

    url = "https://apigetway.track718.net/v2/tracks"
    headers = {"Content-Type": "application/json",
    "Track718-API-Key": f"{cel_api_key}"}

    params = [{"trackNum": track, "code": "gps-truck"}]
    respons = requests.post(url=url, headers=headers, json=params)


    print(respons.status_code)
    print(respons)
    print(respons.json())


def django_api_parcel_info():
    url = 'http://127.0.0.1:8000/api/parcels/'
    response = requests.get(url)
    print(response.status_code)
    print(response.json())


def django_api_post_parcel_history_info():
    url = 'http://164.132.182.145:8000/api_insert_decisions/'
    body = {'parcel_numb': 'CEL8000790773CD', 'time': '2024-01-30 23:11:19',
            'status_name': 'Отказ в выпуске товаров', 'place': '', 'comment': ''}
    response = requests.post(url, json=body)
    print(response.status_code)
    print(response.json())

def test_time():
    # record start time
    start = time.time()

    # define a sample code segment
    a = 0
    for i in range(1000):
        a += (i ** 100)

    # record end time
    end = time.time()

    # print the difference between start
    # and end time in milli. secs
    print("The time of execution of above program is :",
          (end - start) * 10 ** 3, "ms")

def read_json():
    with open("ParcelsProductsData.json", encoding='utf-8') as f:
        data = json.loads(f.read())
        list_all = []
        for line in data:
            print(line)
            for good in line['products']:
                dict_item = {}
                parcel_numb = line["trackingNumber"]
                dict_item["parcel_numb"] = parcel_numb
                description = good['originalName']
                dict_item["description"] = description
                description2 = good['name']
                dict_item["description2"] = description2
                quantity = good['quantity']
                dict_item["quantity"] = quantity
                price = good['price']
                dict_item["price"] = price
                url = good['url']
                dict_item["url"] = url
                weight = good['weight']
                dict_item["weight"] = weight
                tnvedCode = good['tnvedCode']
                dict_item["tnvedCode"] = tnvedCode
                print(parcel_numb)
                list_all.append(dict_item)
        print(list_all)
        df = pd.DataFrame(list_all)
        print(df)
    writer = pd.ExcelWriter(f'РЕЭКСПОРТ Проверка на запрет Декабрь - Январь.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()



def read_sql():
    con = sl.connect('BAZA.db')
    df = pd.read_excel('изменить описания ГБС.xlsx', usecols='A, B')
    df_group = df.groupby('parcel_numb', sort=False)['goods'].agg(','.join)
    df = pd.merge(df, df_group, how='left', left_on='parcel_numb', right_on='parcel_numb')
    df = df.drop_duplicates(subset=['parcel_numb'], keep='first')
    df['goods'] = df['goods_y']
    print(df)
    with con:
        for index, row in df.iterrows():
            parcel_numb = row['parcel_numb']
            descript = row['goods']
            print(parcel_numb, descript)
            query = f"UPDATE baza SET goods = '{descript}' where parcel_numb = '{parcel_numb}'"
            con.execute(query)

#api_track718_add_track("14000222380")
#api_track718("14000222380") # 14000218073
#Scarif_request_few_parcel()
#django_api_post_parcel_history_info()
#track = "14000030437"
#api_track718(track)


#django_api_post_parcel_history_info()
#read_json()

#read_sql()

def GBS_request_events_df():
    df = pd.read_excel('11.xlsx')
    list_chanks = list(chunks(df['parcel_numb'], 25))
    #print(list_chanks)
    n = 0
    dict_of_result = {}
    for chank in list_chanks:
        list_of_dicts = []
        for i in chank:
            json_parc = {"HWBRefNumber": str(i)}
            list_of_dicts.append(json_parc)
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
        try:
            response = requests.post(url=url, json=body,  # http://164.132.182.145:5001
                                     headers=headers)

            json_events = response.json()
            result = json_events['result']
            print(result)
            for parcel in result:
                parcel_numb = parcel['HWBRefNumber']
                print(parcel_numb)
                for event in parcel['events']:
                    comment = event['event_comment']
                    print(comment)
                    dict_of_result[comment] = parcel_numb
            n += 25
            print(n)
        except Exception as e:
            print(e)
            time.sleep(3)
            pass
    print(dict_of_result)
    df = pd.DataFrame(dict_of_result.items(), columns=['Номер реестра', 'Трек'])
    print(df)
    writer = pd.ExcelWriter('GBS_json.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()




def clean_baza():
    con = sl.connect('BAZA.db')
    df = pd.read_excel('ALL4-1.xlsx', header=None)
    i = 0
    for parcel_numb in df[0]:
        i += 1
        print(i)
        with con:
            query = """DELETE FROM baza WHERE parcel_numb = ?
                     AND custom_status_short = ?
                     AND refuse_reason = ?"""

            con.execute(query, (parcel_numb, 'ВЫПУСК', '10716050'))

    query2 = """VACUUM"""
    con.execute(query2)


#clean_baza()

#api_track718_add_track("14000222539")   19.03.2024 05:22
#api_track718("14000222539")


def Scarif_request_events_toxl():
    filename = filedialog.askopenfilename()
    map_scarif_status = {'exds10': 'выпуск товаров без уплаты таможенных платежей',
                         'exds30': 'выпуск возвращаемых товаров разрешен',
                        'exds31': 'требуется уплата таможенных платежей',
                         'exds32': 'выпуск товаров разрешен, таможенные платежи уплачены',
                         'exds33': 'выпуск разрешён, ожидание по временному ввозу',
                      'exds40': 'разрешение на отзыв',
                      'exds70': 'продление срока выпуска',
                      'exds90': 'отказ в выпуске товаров',
                         'exds0': 'статус не определен'}
    df = pd.read_excel(filename, sheet_name=0, engine='openpyxl', usecols='A, AO', header=None)
    dict_of_result = {}
    n = 0
    for parcel_numb in df[0]:
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

            registryNumber = json_events['registryNumber']
            print(parcel_numb)
            print(registryNumber)
            dict_of_result[parcel_numb] = registryNumber
            n += 1
            print(n)
        except Exception as e:
            print(e)


    print(dict_of_result)
    df = pd.DataFrame(dict_of_result.items(), columns=['Номер реестра', 'Трек'])
    print(df)
    writer = pd.ExcelWriter('Scarif_json.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()


Scarif_request_events_toxl()
