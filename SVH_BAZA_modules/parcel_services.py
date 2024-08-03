import datetime
import traceback

from flask import request
from flask import flash
import pandas as pd
import sqlite3 as sl
import numpy as np
import winsound
from .services import insert_user_action, get_user_name


def get_parcel_info_service(done_parcels):

    audiofile = 'None'
    parcel_numb = request.form['parcel_numb']
    parcel_numb = parcel_numb.replace('[CDK]', '')
    print(parcel_numb)
    try:
        con = sl.connect('BAZA.db')
        with con:
            df_parc_events = pd.read_sql(f"SELECT parcel_numb, parcel_plomb_numb, custom_status, "
                                         f"custom_status_short, VH_status, refuse_reason FROM baza where parcel_numb = '{parcel_numb}'", con)
            parcel_plomb_numb = df_parc_events['parcel_plomb_numb'].values[0]
            print(parcel_plomb_numb)
            df_parcel_plomb_info = pd.read_sql(f"SELECT * FROM baza where parcel_plomb_numb = '{parcel_plomb_numb}'",
                                               con)
            try:
                if 'Требуется' in df_parc_events['custom_status'].values[0] or 'уплат' in df_parc_events['refuse_reason'].values[0] or 'Не уплачены' in df_parc_events['refuse_reason'].values[0]:
                    pay_trigger = 'ПЛАТНАЯ'
                else:
                    pay_trigger = ''
            except:
                pay_trigger = ''
            df_parc_quont = len(df_parcel_plomb_info)
            df_parcel_plomb_refuse_info = df_parcel_plomb_info.loc[
                df_parcel_plomb_info['custom_status_short'] == 'ИЗЪЯТИЕ']
            df_parc_refuse_quont = len(df_parcel_plomb_refuse_info)
            df_parcel_plomb_refuse_info['№1'] = np.arange(len(df_parcel_plomb_refuse_info))[::+1] + 1
            df_parcel_plomb_refuse_info = df_parcel_plomb_refuse_info[
                ['№1', 'parcel_numb', 'custom_status_short', 'parcel_plomb_numb',
                 'VH_status', 'parcel_weight', 'goods']]
            df_parcel_plomb_refuse_info = df_parcel_plomb_refuse_info.rename(
                columns={'№1': '№', 'parcel_numb': 'Трек-номер', 'custom_status_short': 'Статус',
                         'parcel_plomb_numb': 'Пломба', 'VH_status': 'ВХ', 'parcel_weight': 'вес', 'goods': 'Товары'})
            df_parcel_plomb_refuse_info['Товары'] = df_parcel_plomb_refuse_info['Товары'].str.slice(0, 200)
            if parcel_plomb_numb == '':
                df_parcel_plomb_refuse_info = pd.DataFrame()
            df_parc_events['№'] = np.arange(len(df_parc_events))[::+1] + 1
            df_parc_events = df_parc_events[['parcel_numb', 'custom_status_short', 'parcel_plomb_numb', 'VH_status', 'custom_status']]
            df_parc_events = df_parc_events.rename(
                columns={'parcel_numb': 'Трек-номер', 'custom_status_short': 'Статус',
                         'parcel_plomb_numb': 'Пломба', 'VH_status': 'ВХ'})
            if df_parc_events.empty:
                flash(f'Посылка не найдена!', category='error')
                #winsound.PlaySound('Snd\Snd_Parcel_Not_Found.wav', winsound.SND_FILENAME)
                audiofile = 'Snd_Parcel_Not_Found.wav'
            elif 'На ВХ' in df_parc_events['ВХ'].values:
                flash(f'Уже размещено', category='success')
            if df_parc_events.loc[df_parc_events['Статус'] == 'ИЗЪЯТИЕ'].empty:
                flash(f'ВЫПУСК')
            # if df_parc_events['custom_status'].values[0] is None:
            #     print(df_parc_events['custom_status'])
            #     flash(f'ИЗЪЯТИЕ на склад {pay_trigger}', category='error')
            #     audiofile = 'Snd_CancelIssue.wav'
            else:
                flash(f'ИЗЪЯТИЕ на склад {pay_trigger}', category='error')
                #winsound.PlaySound('Snd\Snd_CancelIssue.wav', winsound.SND_FILENAME)
                audiofile = 'Snd_CancelIssue.wav'
            done_parcels = done_parcels.append(df_parc_events).drop_duplicates(subset=['Трек-номер'], keep='first')
            done_parcels['№'] = np.arange(len(done_parcels))[::+1] + 1
            done_parcels_plomb_info = done_parcels.loc[done_parcels['Пломба'] == parcel_plomb_numb]
            if len(done_parcels_plomb_info) == df_parc_quont:
                flash(f'Место завершено', category='success')
                winsound.PlaySound('Snd\se_mesta_naid.wav', winsound.SND_FILENAME)
            done_parcels_styl = done_parcels.reset_index()
            done_parcels_styl = done_parcels_styl.drop('index', axis=1)
            done_parcels_styl = done_parcels_styl[
                ['№', 'Трек-номер', 'Статус', 'Пломба', 'ВХ']].drop_duplicates(subset=['Трек-номер'], keep='first')

            done_parcels_styl.fillna("", inplace=True)
            df_parc_events.fillna("", inplace=True)
            df_parcel_plomb_refuse_info.fillna("", inplace=True)
            if len(df_parcel_plomb_refuse_info) == len(done_parcels_styl.loc[(
                    (done_parcels_styl.Статус == 'ИЗЪЯТИЕ') & (done_parcels_styl.Пломба == parcel_plomb_numb))]):
                flash(f'Все отказы найдены!', category='success')

            def highlight_last_row_2(done_parcels_styl):
                return ['background-color: #FF0000' if 'ИЗЪЯТИЕ' in str(i) else '' for i in done_parcels_styl]

            if 'ИЗЪЯТИЕ' in done_parcels_styl['Статус'].values:
                done_parcels_styl = done_parcels_styl.style.apply(highlight_last_row_2).hide_index()

        object_name = parcel_numb
        comment = 'Отбор посылок: Просмотрена посылка'
        insert_user_action(object_name, comment)
    except Exception as e:
        print(str(traceback.format_exc()))

    return (df_parcel_plomb_refuse_info, done_parcels_styl, parcel_numb, audiofile,
     df_parc_quont, df_parc_refuse_quont, done_parcels, parcel_plomb_numb)


def clean_working_place_service(done_parcels):
    print(done_parcels)
    done_parcels_VH = done_parcels.loc[done_parcels['Статус'] == 'ИЗЪЯТИЕ']
    con = sl.connect('BAZA.db')
    with con:
        for parcel_numb in done_parcels_VH['Трек-номер']:
            con.execute(
                f"Update baza set VH_status = 'На ВХ', parcel_plomb_numb = ''  where parcel_numb = '{parcel_numb}'")

    parcel_plomb_numb = done_parcels_VH['Пломба'].values[0]  # done_parcels
    done_parcels_VH['parcel_plomb_numb'] = done_parcels_VH['Пломба']
    df_plombs_to_sql = done_parcels_VH['parcel_plomb_numb'].drop_duplicates()
    with con:
        df_plombs_to_sql.to_sql('plombs_with_pullings', con, if_exists='append', index=False)
    done_parcels = pd.DataFrame()
    return done_parcels, parcel_plomb_numb


def search_parcel_sql_service():
    now = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    audiofile = 'None'
    user_name, user_id = get_user_name()
    print(user_name, user_id)
    try:
        parcel_numb = request.form['parcel_numb']
        parcel_numb = parcel_numb.replace('[CDK]', '')
        print(parcel_numb)
        try:
            con = sl.connect('BAZA.db')
            with con:
                con.execute(
                    f"Update parcels_refuses set parcel_find_status = 'НАЙДЕНА' where parcel_numb = '{parcel_numb}'")
                parcel_plomb_numb = pd.read_sql(
                    f"SELECT parcel_plomb_numb FROM parcels_refuses where parcel_numb = '{parcel_numb}'",
                    con).values[0][0]
                print(parcel_plomb_numb)
                df_plomb_numb = pd.read_sql(
                    f"SELECT user_id, parcel_numb, parcel_plomb_numb, custom_status_short, parcel_find_status, "
                    f"goods, parcel_weight FROM parcels_refuses where parcel_plomb_numb = '{parcel_plomb_numb}'",
                    con)
                df_refuses = df_plomb_numb.loc[df_plomb_numb['custom_status_short'] == 'ИЗЪЯТИЕ'].fillna('')
                df_parcel = df_plomb_numb.loc[df_plomb_numb['parcel_numb'] == parcel_numb]
                user_id_in_baze = df_plomb_numb['user_id'].values[0]
                print(user_id_in_baze)
                print(user_id_in_baze)
                if str(user_id) == str(user_id_in_baze):
                    con.execute(
                        f"Update parcels_refuses set parcel_find_status = 'НАЙДЕНА', time = '{now}' where parcel_numb = '{parcel_numb}'")
                    custom_status_short = df_parcel['custom_status_short'].values[0]
                    print(custom_status_short)
                    print(parcel_plomb_numb)
                    if custom_status_short == "ВЫПУСК":
                        flash(f'ВЫПУСК', category='success')
                        audiofile = 'zvuk-vezeniya.wav'
                    else:
                        flash(f'ИЗЪЯТИЕ', category='error')
                        audiofile = 'Snd_CancelIssue.wav'
                    if df_refuses.empty:
                        flash(f'Все место выпущено!', category='success')
                        audiofile = 'Snd_All_Issue.wav'
                    df_refuses = df_refuses[['parcel_numb', 'parcel_find_status', 'custom_status_short',
                                             'parcel_weight', 'goods', 'user_id']].sort_values('parcel_find_status',
                                 ascending=True)

                    qt_refuse = len(df_refuses)
                    qt_all = len(df_plomb_numb)
                    qt_found = len(df_refuses.loc[df_refuses['parcel_find_status'] == 'НАЙДЕНА'])
                    df_refuses['№'] = np.arange(len(df_refuses))[::+1] + 1
                    df_refuses = df_refuses[['№', 'parcel_numb', 'parcel_find_status', 'custom_status_short',
                                             'parcel_weight', 'goods', 'user_id']]
                    df_refuses.columns = ['№', 'Трек', 'Скан', 'ТО (кратко)', 'Вес', 'Товары', 'user_id']
                    if qt_refuse == qt_found:
                        flash(f'Все отказы найдены!', category='success')

                else:
                    flash(f'Не совпадает user_id', category='error')
                    audiofile = 'Snd_CancelIssue.wav'
                object_name = parcel_numb
                comment = 'Отбор посылок: Просмотрена посылка'
                insert_user_action(object_name, comment)
        except Exception as e:
            print(str(traceback.format_exc()))
    except:
        parcel_numb = None
        df_refuses = None
    return parcel_numb, parcel_plomb_numb, audiofile, df_refuses, qt_refuse, qt_all, qt_found


def clean_working_place_sql_service(parcel_plomb_numb):
    user_name, user_id = get_user_name()
    con = sl.connect('BAZA.db')
    with con:
        df_refuses = pd.read_sql(
            f"SELECT user_id, parcel_numb, parcel_plomb_numb, custom_status_short, parcel_find_status, goods FROM parcels_refuses where parcel_plomb_numb = '{parcel_plomb_numb}' and custom_status_short = 'ИЗЪЯТИЕ'",
            con)
        user_id_in_baze = df_refuses['user_id'].values[0]
        if str(user_id_in_baze) == str(user_id):
            for parcel in df_refuses['parcel_numb']:
                parcel_find_status = pd.read_sql(
                    f"SELECT parcel_find_status FROM parcels_refuses where parcel_numb = '{parcel}'",
                    con).values[0]
                if parcel_find_status == 'НАЙДЕНА':
                    con.execute(
                        f"Update baza set VH_status = 'На ВХ', parcel_plomb_numb = ''  where parcel_numb = '{parcel}'")
                else:
                    con.execute(
                        f"Update baza set VH_status = 'НЕ НАЙДЕНА', parcel_plomb_numb = ''  where parcel_numb = '{parcel}'")
                    flash(f'Посылка не была просканированна: {parcel}, проставлен статус "НЕ НАЙДЕНА"', category='error')
            df_plombs_to_sql = df_refuses['parcel_plomb_numb'].drop_duplicates()

            with con:
                df_plombs_to_sql.to_sql('plombs_with_pullings', con, if_exists='append', index=False)
            return df_refuses


def add_to_zone_service():
    user_name, user_id = get_user_name()
    audiofile = 'None'
    try:
        parcel_numb = request.form['parcel_numb']
        parcel_numb = parcel_numb.replace('[CDK]', '')
        con_vec = sl.connect('VECTORS.db')
        try:
            vector = con_vec.execute(
                f"SELECT vector FROM VECTORS where parcel_numb = '{parcel_numb}'").fetchone()[
                0]
            flash(f'{vector}!', category='success')
        except Exception as e:
            vector = None
            print(e)
        con = sl.connect('BAZA.db')
        with con:
            df_parcel = pd.read_sql(
                f"SELECT parcel_numb, custom_status_short, zone FROM baza where parcel_numb = '{parcel_numb}'",
                con)
            zone = df_parcel['zone'].values[0]
            if not df_parcel.empty:
                try:
                    con.execute(
                        f"INSERT OR REPLACE INTO add_to_zone ('parcel_numb', 'zone', 'user_id') VALUES('{parcel_numb}', "
                        f"'{zone}', '{user_id}')")
                except sl.IntegrityError:
                    print('cant insert')
                custom_status_short = df_parcel['custom_status_short'].values[0]
                if custom_status_short == "ВЫПУСК":
                    flash(f'ВЫПУСК', category='success')
                    audiofile = 'Snd_Issue.wav'
                else:
                    flash(f'ИЗЪЯТИЕ', category='error')
                    audiofile = 'Snd_CancelIssue.wav'
                object_name = parcel_numb
                comment = f'Посылка отмечена для добавления на Зону хранения'
                insert_user_action(object_name, comment)
            else:
                flash(f'Посылка не найдена!', category='error')
                audiofile = 'Snd_CancelIssue.wav'
            df_user_work = pd.read_sql(f"SELECT * FROM add_to_zone where user_id = '{user_id}'",
            con).fillna('').sort_values(by='ID', ascending=False)

            df_user_work['№'] = np.arange(len(df_user_work))[::-1] + 1
            df_user_work = df_user_work[['№', 'parcel_numb', 'zone', 'user_id']]

    except:
        print(str(traceback.format_exc()))
        parcel_numb = None
        vector = None
        audiofile = 'Snd_CancelIssue.wav'
    return parcel_numb, vector, audiofile, df_user_work


def add_to_place_sql_service():
    user_name, user_id = get_user_name()
    audiofile = 'None'
    try:
        parcel_numb = request.form['parcel_numb']
        parcel_numb = parcel_numb.replace('[CDK]', '')
        con_vec = sl.connect('VECTORS.db')
        try:
            vector = con_vec.execute(
                f"SELECT vector FROM VECTORS where parcel_numb = '{parcel_numb}'").fetchone()[
                0]
            flash(f'{vector}!', category='success')
        except Exception as e:
            vector = None
            print(e)
        con = sl.connect('BAZA.db')
        with con:
            df_parcel = pd.read_sql(
                f"SELECT parcel_numb, parcel_plomb_numb, custom_status_short FROM baza where parcel_numb = '{parcel_numb}'",
                con)
            if not df_parcel.empty:
                custom_status_short = df_parcel['custom_status_short'].values[0]
                parcel_plomb_numb = df_parcel['parcel_plomb_numb'].values[0]
                try:
                    con.execute(
                        f"INSERT OR REPLACE INTO add_to_place ('parcel_numb', 'parcel_plomb_numb',"
                        f"'custom_status_short', 'user_id') VALUES('{parcel_numb}', '{parcel_plomb_numb}'"
                        f", '{custom_status_short}', '{user_id}')")
                except sl.IntegrityError:
                    print('cant insert')
                custom_status_short = df_parcel['custom_status_short'].values[0]
                if custom_status_short == "ВЫПУСК":
                    flash(f'ВЫПУСК', category='success')
                else:
                    flash(f'ИЗЪЯТИЕ', category='error')
                    audiofile = 'Snd_CancelIssue.wav'
                object_name = parcel_numb
                comment = f'Посылка отмечена для добавления под пломбу'
                insert_user_action(object_name, comment)
            else:
                flash(f'Посылка не найдена!', category='error')
                audiofile = 'Snd_CancelIssue.wav'
            df_user_work = pd.read_sql(f"SELECT * FROM add_to_place where user_id = '{user_id}'",
            con).fillna('').sort_values(by='ID', ascending=False)

            df_user_work['№'] = np.arange(len(df_user_work))[::-1] + 1
            df_user_work = df_user_work[['№', 'parcel_numb', 'parcel_plomb_numb', 'custom_status_short', 'user_id']]

    except:
        print(str(traceback.format_exc()))
        parcel_numb = None
        vector = None
        audiofile = 'Snd_CancelIssue.wav'
    return parcel_numb, vector, audiofile, df_user_work
