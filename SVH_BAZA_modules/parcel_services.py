from flask import request
from flask import flash
import pandas as pd
import sqlite3 as sl
import numpy as np
import winsound
from .services import insert_user_action


def get_parcel_info_service(done_parcels):
    audiofile = 'None'
    parcel_numb = request.form['parcel_numb']
    parcel_numb = parcel_numb.replace('[CDK]', '')
    try:
        con = sl.connect('BAZA.db')
        with (con):
            df_parc_events = pd.read_sql(f"SELECT * FROM baza where parcel_numb = '{parcel_numb}'", con)
            parcel_plomb_numb = df_parc_events['parcel_plomb_numb'].values[0]
            df_parcel_plomb_info = pd.read_sql(f"SELECT * FROM baza where parcel_plomb_numb = '{parcel_plomb_numb}'",
                                               con)
            if 'Требуется' in df_parc_events['custom_status'].values[0] or 'уплат' in df_parc_events['refuse_reason'].values[0] or 'Не уплачены' in df_parc_events['refuse_reason'].values[0]:
                pay_trigger = 'ПЛАТНАЯ'
            else:
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
            df_parc_events = df_parc_events[['parcel_numb', 'custom_status_short', 'parcel_plomb_numb', 'VH_status']]
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
        print(e)

    return (df_parcel_plomb_refuse_info, done_parcels_styl, parcel_numb, audiofile,
     df_parc_quont, df_parc_refuse_quont, done_parcels)


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

