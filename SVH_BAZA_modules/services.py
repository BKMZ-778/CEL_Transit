from flask import request
import sqlite3 as sl
import datetime
import os
import logging


download_folder = 'C:/Users/User/Desktop/ДОКУМЕНТЫ/'
download_folder_allmanif = 'C:/Users/User/Desktop/ДОКУМЕНТЫ/ОТГРУЖЕННОЕ'
addition_folder = f'{download_folder}Места-Паллеты/'
if not os.path.isdir(download_folder):
    os.makedirs(download_folder, exist_ok=True)
if not os.path.isdir(addition_folder):
    os.makedirs(addition_folder, exist_ok=True)
if not os.path.isdir(download_folder_allmanif):
    os.makedirs(download_folder_allmanif, exist_ok=True)

style = ('<style>.dataframe th{background: rgb(255,255,255);background: radial-gradient(circle, rgba(255,255,255,'
         '1) 0%, rgba(236,236,236,1) 100%);padding: 5px;color: #343434;font-family: monospace;font-size: '
         '110%;border:2px solid #e0e0e0;text-align:left !important;}.dataframe{border: 3px solid #ffebeb '
         '!important;}</style>')


map_eng_to_rus = {'registration_numb': 'Реестр', 'party_numb': 'Партия',
                                    'parcel_numb': 'Трек-номер', 'parcel_plomb_numb': 'Пломба', 'parcel_weight': 'вес',
                                    'custom_status': 'Статус ТО', 'custom_status_short': 'Статус ТО (кратк)',
                                    'decision_date': 'Дата решения',
                                    'refuse_reason': 'Причина отказа',
                                    'pallet': 'Паллет',
                                    'zone': 'Зона',
                                  'VH_status': 'Статус ВХ',
                                  'goods': 'Товары'}

def create_databases():
    con = sl.connect('BAZA-reports.db')
    with con:
        baza = con.execute("select count(*) from sqlite_master where type='table' and name='report'")
        for row in baza:
            # если таких таблиц нет
            if row[0] == 0:
                # создаём таблицу
                with con:
                    con.execute("""
                                CREATE TABLE report (
                                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                user_id INTEGER NOT NULL,
                                object_name VARCHAR(30),
                                comment VARCHAR(120),
                                action_date DATETIME
                                );
                            """)

    con = sl.connect('BAZA.db')
    with con:
        baza = con.execute("select count(*) from sqlite_master where type='table' and name='df_plomb_to_manifest'")
        for row in baza:
            # если таких таблиц нет
            if row[0] == 0:
                # создаём таблицу
                with con:
                    con.execute("""
                                                            CREATE TABLE df_plomb_to_manifest (
                                                            ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                            parcel_plomb_numb VARCHAR(20)
                                                            );
                                                        """)
    con.commit()
    con.close()

    con = sl.connect('BAZA.db')
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
    con.commit()
    con.close()

def setup_logger(name, log_file, level=logging.INFO):
    logging.basicConfig(format=u'%(levelname)-8s [%(asctime)s] %(message)s')  # filename=u'mylog.log'
    handler = logging.FileHandler(log_file)
    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.addHandler(handler)
    return logger

logger = setup_logger('logger', 'mylog.log')

logger_change_plob = setup_logger('logger_change_plob', f'ИЗМЕНЕНИЯ ПЛОМБ.log')
logger_API = setup_logger('logger_API', f'API_insert.log')


def get_user_name():
    try:
        user_id = request.cookies.get('user_id')
        con_user = sl.connect("db.sqlite")
        query = f"Select name from users where id = {user_id}"
        user_name = con_user.execute(query).fetchone()[0]
    except:
        user_id = ''
        user_name = 'нет авторизации'
    return user_name, user_id



def insert_user_action(object_name, comment):
    try:
        user_id = request.cookies.get('user_id')
    except:
        user_id = "API"
    if user_id is None:
        user_id = "API"
    print(user_id)
    now_time = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    con = sl.connect('BAZA-reports.db')
    with con:
        cur = con.cursor()
        action_date = now_time
        statement = "INSERT INTO report (user_id, object_name, comment, action_date) VALUES (?, ?, ?, ?)"
        cur.execute(statement, [user_id, object_name, comment, action_date])
        con.commit()


