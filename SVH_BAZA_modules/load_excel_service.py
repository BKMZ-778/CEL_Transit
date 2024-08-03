from flask import Flask, jsonify, request, render_template, redirect, url_for, send_file
from flask import abort
from flask import flash
import pandas as pd
import sqlite3 as sl
import os
import winsound
from sqlalchemy import text, create_engine

from SVH_BAZA_modules.services import logger


db_url = "mysql+mysqlconnector://{USER}:{PWD}@{HOST}/{DBNAME}"
db_url = db_url.format(
    USER="root",
    PWD="jPouKY2zy3R6",
    HOST="localhost",
    DBNAME="baza",
    auth_plugin='mysql_native_password'
)
engine = create_engine(db_url, echo=False)


def load_sample_manifest_service(uploaded_file, filename):
    uploaded_file.save(uploaded_file.filename)
    df_track = pd.read_excel(filename, sheet_name=0, engine='openpyxl', usecols='A, AO')
    df_track = df_track[['Номер отправления ИМ', 'Номер накладной СДЭК']]
    df_track = df_track.rename(columns={'Номер отправления ИМ': 'track_numb',
                            'Номер накладной СДЭК': 'parcel_numb'})
    print(df_track)
    if df_track['track_numb'].astype(str).str.contains('#н/д').any() or df_track['track_numb'].isnull().any():
        flash(f'Ошибка загрузки треков: В колонке Треков есть пустые значения или #н/д, поправьте и загрузите заново', category='error')
    else:
        con_track = sl.connect('TRACKS.db')
        with con_track:
            data = con_track.execute("select count(*) from sqlite_master where type='table' and name='tracks'")
            for row in data:
                # если таких таблиц нет
                if row[0] == 0:
                    # создаём таблицу
                    with con_track:
                        con_track.execute("""
                                                CREATE TABLE tracks (
                                                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                parcel_numb VARCHAR(25) NOT NULL UNIQUE ON CONFLICT REPLACE,
                                                track_numb VARCHAR(25)
                                                );
                                                """)
            df_track.to_sql('tracks', con=con_track, if_exists='append', index=False)
            con_track.commit()
            print(df_track)
    df = pd.read_excel(filename, sheet_name=0, engine='openpyxl')
    df = df[['Номер накладной СДЭК', 'Номер пломбы', 'Наименование товара', '№ AWB', 'Общий Вес места (накладной)']]
    df = df.rename(columns={'Номер накладной СДЭК': 'parcel_numb',
                           'Номер пломбы': 'parcel_plomb_numb',
                           'Наименование товара': 'goods',
                           '№ AWB': 'party_numb', 'Общий Вес места (накладной)': 'parcel_weight'})
    logger.warning(df)
    df_group = df.groupby('parcel_numb', sort=False)['goods'].agg(','.join)
    logger.warning(df_group)
    df = pd.merge(df, df_group, how='left', left_on='parcel_numb', right_on='parcel_numb')
    df = df.drop_duplicates(subset=['parcel_numb'], keep='first')
    df['goods'] = df['goods_y']
    df = df[['parcel_numb', 'parcel_plomb_numb', 'goods', 'party_numb', 'parcel_weight']]
    df['custom_status'] = 'Unknown'
    df['custom_status_short'] = 'ИЗЪЯТИЕ'
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
        party_numb = df['party_numb'].values[0]
        party_numb_isalready_in = pd.read_sql(f"Select party_numb from baza where party_numb = '{party_numb}'", con)
        if party_numb_isalready_in.empty:
            df.to_sql('baza', con=con, if_exists='append', index=False)
        else:
            for parcel_numb in df['parcel_numb']:
                row_isalready_in = pd.read_sql(f"Select * from baza where parcel_numb = '{parcel_numb}'", con)
                row = df.loc[df['parcel_numb'] == parcel_numb]
                if row_isalready_in.empty:
                    df_to_append = df_to_append.append(row)
                    logger.warning(row)
                else:
                    goods = df.loc[df['parcel_numb'] == parcel_numb]['goods'].values[0]
                    parcel_weight = df.loc[df['parcel_numb'] == parcel_numb]['parcel_weight'].values[0]
                    logger.warning(goods)

                    party_numb_in_base = row_isalready_in['party_numb'].values[0]
                    if party_numb_in_base is None:
                        party_numb = df.loc[df['parcel_numb'] == parcel_numb]['party_numb'].values[0]
                        parcel_plomb_numb = str(df.loc[df['parcel_numb'] == parcel_numb]['parcel_plomb_numb'].values[0])
                        parcel_weight = df.loc[df['parcel_numb'] == parcel_numb]['parcel_weight'].values[0]
                        con.execute("Update baza set goods = ?, "
                                    "party_numb = ?, "
                                    "parcel_plomb_numb = ?,"
                                    "parcel_weight = ?"
                                    "where parcel_numb = ?",
                                    (goods, party_numb, parcel_plomb_numb, parcel_weight, parcel_numb))
                    else:
                        con.execute("Update baza set party_numb = ?, goods = ?, parcel_weight = ? where parcel_numb = ?",
                                    (party_numb, goods, parcel_weight, parcel_numb))
            logger.warning(df_to_append)
            df_to_append.to_sql('baza', con=con, if_exists='append', index=False)
        try:
            df_vector = pd.read_excel(filename, sheet_name=0, engine='openpyxl')
            df_vector = df_vector[['№ AWB', 'Номер пломбы', 'Номер отправления ИМ', 'Примечание']]
            df_vector = df_vector.rename(columns={'№ AWB': 'party_numb', 'Номер пломбы': 'parcel_plomb_numb', 'Номер отправления ИМ': 'parcel_numb',
                                                  'Примечание': 'vector'})
            con = sl.connect('VECTORS.db')
            with con:
                baza = con.execute("select count(*) from sqlite_master where type='table' and name='vectors'")
                for row in baza:
                    # если таких таблиц нет
                    if row[0] == 0:
                        # создаём таблицу
                        with con:
                            con.execute("""
                                                        CREATE TABLE vectors (
                                                        ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                        party_numb VARCHAR(20),
                                                        parcel_plomb_numb VARCHAR(20),
                                                        parcel_numb VARCHAR(20) NOT NULL UNIQUE ON CONFLICT REPLACE,
                                                        vector
                                                        );
                                                    """)
                df_vector.to_sql('vectors', con=con, if_exists='append', index=False)
            flash(f'Инфо по последней миле загружена')
        except Exception as e:
            flash(f'Нет инфо по последней миле {e}')

        flash(f'Шаблон загружен')
        winsound.PlaySound('Snd\sample_load.wav', winsound.SND_FILENAME)


def load_decisions_service(uploaded_file, filename):
    uploaded_file.save(uploaded_file.filename)
    df = pd.read_excel(filename, sheet_name=0, engine='openpyxl')
    df['Вес брутто'] = df['Вес брутто'].replace(',', '.', regex=True).astype(float)
    df['Статус_ТО'] = df['Статус ТО']
    df['Статус_ТО'] = df['Статус_ТО'].replace(to_replace='10-в',
                                              value='В', regex=True)
    df['Статус_ТО'] = df['Статус_ТО'].replace(to_replace='32-в',
                                              value='В', regex=True)
    df['Статус_ТО'] = df['Статус_ТО'].replace(to_replace='Выпуск товаров без уплаты таможенных платежей', value='ВЫПУСК', regex=True)
    df['Статус_ТО'] = df['Статус_ТО'].replace(to_replace='Выпуск товаров разрешен, таможенные платежи уплачены', value='ВЫПУСК', regex=True)
    for cel in df['Статус_ТО']:
        if cel != 'ВЫПУСК':
            df.loc[df['Статус_ТО'] == cel, 'Статус_ТО'] = 'ИЗЪЯТИЕ'
    df['Пломба'] = df['Пломба'].astype(str)

    df = df.rename(columns={'Рег. номер': 'registration_numb', 'Номер общей накладной': 'party_numb',
                            'Трек-номер': 'parcel_numb', 'Пломба': 'parcel_plomb_numb', 'Вес брутто': 'parcel_weight',
                            'Статус ТО': 'custom_status', 'Статус_ТО': 'custom_status_short', 'Дата решения': 'decision_date',
                            'Причина отказа ТО': 'refuse_reason'})
    df = df.drop('Код причины отказа', axis=1)
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
        df_to_append = pd.DataFrame()
        for parcel_numb in df['parcel_numb']:
            row_isalready_in = pd.read_sql(f"Select * from baza where parcel_numb = '{parcel_numb}'", con)
            row = df.loc[df['parcel_numb'] == parcel_numb]
            if row_isalready_in.empty:
                # row.to_sql('baza', con=con, if_exists='append', index=False)
                df_to_append = df_to_append.append(row)
                logger.warning(row)
            else:
                logger.warning(row)
                custom_status_short = df.loc[df['parcel_numb'] == parcel_numb]['custom_status_short'].values[0]
                logger.warning(custom_status_short)
                custom_status = df.loc[df['parcel_numb'] == parcel_numb]['custom_status'].values[0]
                logger.warning(custom_status)
                registration_numb = df.loc[df['parcel_numb'] == parcel_numb]['registration_numb'].values[0]
                logger.warning(registration_numb)
                parcel_weight = df.loc[df['parcel_numb'] == parcel_numb]['parcel_weight'].values[0]
                decision_date = df.loc[df['parcel_numb'] == parcel_numb]['decision_date'].values[0]
                refuse_reason = df.loc[df['parcel_numb'] == parcel_numb]['refuse_reason'].values[0]

                con.execute(f"Update baza set registration_numb = '{registration_numb}',"
                            f" custom_status = '{custom_status}',"
                            f" custom_status_short = '{custom_status_short}',"
                            f" parcel_weight = '{parcel_weight}',"
                            f" decision_date = '{decision_date}',"
                            f" refuse_reason = '{refuse_reason}' where parcel_numb = '{parcel_numb}'")
                row_isalready_in = pd.read_sql(f"Select * from baza where parcel_numb = '{parcel_numb}'", con)
                logger.warning(row_isalready_in)
            # df_isalready_in = df_isalready_in.append(row_isalready_in)
        logger.warning(df_to_append)
        print(df_to_append)
        writer = pd.ExcelWriter(f'df_to_append.xlsx', engine='xlsxwriter')
        df_to_append.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()
        df_to_append.to_sql('baza', con=con, if_exists='append', index=False)
        flash(f'Решения загружены')
        winsound.PlaySound('Snd\esheniya_zagruzhenu.wav', winsound.SND_FILENAME)


# def load_sample_manifest_service_mysql(uploaded_file, filename):
#     uploaded_file.save(uploaded_file.filename)
#     df_track = pd.read_excel(filename, sheet_name=0, engine='openpyxl', usecols='A, AO')
#     df_track = df_track[['Номер отправления ИМ', 'Номер накладной СДЭК']]
#     df_track = df_track.rename(columns={'Номер отправления ИМ': 'parcel_numb',
#                             'Номер накладной СДЭК': 'track_numb'})
#     print(df_track)
#     if df_track['track_numb'].astype(str).str.contains('#н/д').any() or df_track['track_numb'].isnull().any():
#         flash(f'Ошибка загрузки треков: В колонке Треков есть пустые значения или #н/д, поправьте и загрузите заново', category='error')
#     else:
#         con_track = sl.connect('TRACKS.db')
#         with con_track:
#             data = con_track.execute("select count(*) from sqlite_master where type='table' and name='tracks'")
#             for row in data:
#                 # если таких таблиц нет
#                 if row[0] == 0:
#                     # создаём таблицу
#                     with con_track:
#                         con_track.execute("""
#                                                 CREATE TABLE tracks (
#                                                 ID INTEGER PRIMARY KEY AUTOINCREMENT,
#                                                 parcel_numb VARCHAR(25) NOT NULL UNIQUE ON CONFLICT REPLACE,
#                                                 track_numb VARCHAR(25)
#                                                 );
#                                                 """)
#             df_track.to_sql('tracks', con=con_track, if_exists='append', index=False)
#             con_track.commit()
#             print(df_track)
#     df = pd.read_excel(filename, sheet_name=0, engine='openpyxl')
#     df = df[['Номер отправления ИМ', 'Номер пломбы', 'Наименование товара', '№ AWB', 'Общий Вес места (накладной)']]
#     df = df.rename(columns={'Номер отправления ИМ': 'parcel_numb',
#                            'Номер пломбы': 'parcel_plomb_numb',
#                            'Наименование товара': 'goods',
#                            '№ AWB': 'party_numb', 'Общий Вес места (накладной)': 'parcel_weight'})
#     logger.warning(df)
#     df_group = df.groupby('parcel_numb', sort=False)['goods'].agg(','.join)
#     logger.warning(df_group)
#     df = pd.merge(df, df_group, how='left', left_on='parcel_numb', right_on='parcel_numb')
#     df = df.drop_duplicates(subset=['parcel_numb'], keep='first')
#     df['goods'] = df['goods_y']
#     df = df[['parcel_numb', 'parcel_plomb_numb', 'goods', 'party_numb', 'parcel_weight']]
#     df['custom_status'] = 'Unknown'
#     df['custom_status_short'] = 'ИЗЪЯТИЕ'
#     # открываем базу
#     con = engine.connect()
#     with con:
#         try:
#             create_baza_table_query = """
#                                     CREATE TABLE baza (
#                                     ID INTEGER PRIMARY KEY AUTO_INCREMENT,
#                                     registration_numb VARCHAR(25),
#                                     party_numb VARCHAR(20),
#                                     parcel_numb VARCHAR(20) NOT NULL UNIQUE,
#                                     parcel_plomb_numb VARCHAR(20),
#                                     parcel_weight DECIMAL(7,3) NOT NULL,
#                                     custom_status VARCHAR(400),
#                                     custom_status_short VARCHAR(8),
#                                     decision_date VARCHAR(20),
#                                     refuse_reason VARCHAR(400),
#                                     pallet VARCHAR(10),
#                                     zone VARCHAR(10),
#                                     VH_status VARCHAR(10),
#                                     goods VARCHAR(499),
#                                     vector VARCHAR(50),
#                                     track_numb VARCHAR(30)
#                                     );
#                                 """
#             try:
#                 con.execute(text(create_baza_table_query))
#             except Exception as e:
#                 print(e)
#         except Exception as e:
#             print(e)
#         df_to_append = pd.DataFrame()
#         party_numb = df['party_numb'].values[0]
#         party_numb_isalready_in = pd.read_sql(text(f"Select party_numb from baza where party_numb = '{party_numb}'"), con)
#         print(party_numb_isalready_in)
#         df['goods'] = df['goods'].str[:499]
#         with engine.begin() as con:
#             if party_numb_isalready_in.empty:
#                 print(df)
#                 df.to_sql('baza', con=con, if_exists='append', index=False)
#             else:
#                 for parcel_numb in df['parcel_numb']:
#                     row_isalready_in = pd.read_sql(text(f"Select * from baza where parcel_numb = '{parcel_numb}'"), con)
#                     row = df.loc[df['parcel_numb'] == parcel_numb]
#                     if row_isalready_in.empty:
#                         df_to_append = df_to_append.append(row)
#                         logger.warning(row)
#                     else:
#                         goods = df.loc[df['parcel_numb'] == parcel_numb]['goods'].values[0]
#                         parcel_weight = df.loc[df['parcel_numb'] == parcel_numb]['parcel_weight'].values[0]
#                         logger.warning(goods)
#
#                         party_numb_in_base = row_isalready_in['party_numb'].values[0]
#                         if party_numb_in_base is None:
#                             party_numb = df.loc[df['parcel_numb'] == parcel_numb]['party_numb'].values[0]
#                             parcel_plomb_numb = str(df.loc[df['parcel_numb'] == parcel_numb]['parcel_plomb_numb'].values[0])
#                             parcel_weight = df.loc[df['parcel_numb'] == parcel_numb]['parcel_weight'].values[0]
#                             con.execute(text("Update baza set goods = ?, "
#                                         "party_numb = ?, "
#                                         "parcel_plomb_numb = ?,"
#                                         "parcel_weight = ?"
#                                         "where parcel_numb = ?"),
#                                         (goods, party_numb, parcel_plomb_numb, parcel_weight, parcel_numb))
#                         else:
#                             con.execute(text("Update baza set party_numb = ?, goods = ?, parcel_weight = ? where parcel_numb = ?"),
#                                         (party_numb, goods, parcel_weight, parcel_numb))
#                 logger.warning(df_to_append)
#                 df_to_append.to_sql('baza', con=con, if_exists='append', index=False)
#         try:
#             df_vector = pd.read_excel(filename, sheet_name=0, engine='openpyxl')
#             df_vector = df_vector[['№ AWB', 'Номер пломбы', 'Номер отправления ИМ', 'Примечание']]
#             df_vector = df_vector.rename(columns={'№ AWB': 'party_numb', 'Номер пломбы': 'parcel_plomb_numb', 'Номер отправления ИМ': 'parcel_numb',
#                                                   'Примечание': 'vector'})
#             con = sl.connect('VECTORS.db')
#             with con:
#                 baza = con.execute("select count(*) from sqlite_master where type='table' and name='vectors'")
#                 for row in baza:
#                     # если таких таблиц нет
#                     if row[0] == 0:
#                         # создаём таблицу
#                         with con:
#                             con.execute("""
#                                                         CREATE TABLE vectors (
#                                                         ID INTEGER PRIMARY KEY AUTOINCREMENT,
#                                                         party_numb VARCHAR(20),
#                                                         parcel_plomb_numb VARCHAR(20),
#                                                         parcel_numb VARCHAR(20) NOT NULL UNIQUE ON CONFLICT REPLACE,
#                                                         vector
#                                                         );
#                                                     """)
#                 df_vector.to_sql('vectors', con=con, if_exists='append', index=False)
#             flash(f'Инфо по последней миле загружена')
#         except Exception as e:
#             flash(f'Нет инфо по последней миле {e}')
#
#         flash(f'Шаблон загружен')
#         winsound.PlaySound('Snd\sample_load.wav', winsound.SND_FILENAME)