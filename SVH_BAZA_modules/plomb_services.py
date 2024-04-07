from flask import request
from flask import flash
import pandas as pd
import sqlite3 as sl
import numpy as np
import winsound
from .services import insert_user_action


def get_plomb_come_work_service(party_numb):
    parcel_plomb_numb = request.form['parcel_plomb_numb']
    con_vect = sl.connect('VECTORS.db')
    vector = None
    with con_vect:
        try:
            vector = con_vect.execute(
                f"SELECT vector FROM VECTORS where parcel_plomb_numb = '{parcel_plomb_numb}'").fetchone()[0]
            flash(f'{vector}!', category='success')
            print(vector)
        except Exception as e:
            print(e)
    con = sl.connect('BAZA.db')
    df_not_custom = pd.read_excel('неподанные.xlsx', sheet_name=0, engine='openpyxl', header=None)
    if df_not_custom[0].astype(str).str.contains(parcel_plomb_numb).any():
        flash(f'НЕПОДАНА Пломба {parcel_plomb_numb} !!', category='error')
        winsound.PlaySound('Snd\Snd_CancelIssue.wav', winsound.SND_FILENAME)
    with con:
        df_check_plomb = pd.read_sql(
            f"SELECT * FROM plombs where parcel_plomb_numb = '{parcel_plomb_numb}' COLLATE NOCASE", con)
        if df_check_plomb.empty:
            flash(f'Пломба {parcel_plomb_numb} не найдена в партии!', category='error')
            winsound.PlaySound('Snd\Snd_CancelIssue.wav', winsound.SND_FILENAME)
        else:
            df_check_refuse = pd.read_sql(
                f"SELECT parcel_plomb_numb FROM baza where parcel_plomb_numb = '{parcel_plomb_numb}' and custom_status_short = 'ИЗЪЯТИЕ'",
                con)
            if df_check_refuse.empty:
                flash(f'ВЫПУСК МЕСТА', category='success')
                winsound.PlaySound('Snd\zvuk-vezeniya.wav', winsound.SND_FILENAME)
            con.execute(
                f"Update plombs set parcel_plomb_status = 'Принят' where parcel_plomb_numb = '{parcel_plomb_numb}'")
        df = pd.read_sql(f"SELECT * FROM plombs where party_numb = '{party_numb}' COLLATE NOCASE", con)
        df = df.drop_duplicates(subset='parcel_plomb_numb').loc[df['parcel_plomb_numb'] != '']
        df['№'] = np.arange(len(df))[::+1] + 1
        df = df[['№', 'parcel_plomb_numb', 'party_numb', 'parcel_plomb_status']]
        if df.loc[df['parcel_plomb_status'] == 'Ожидаем'].empty:
            flash(f'Все места найдены', category='success')
            winsound.PlaySound('Snd\se_mesta_naid.wav', winsound.SND_FILENAME)
        index = False
        quont_all_plombs = len(df)
        quont_plomb_done = len(df.loc[df['parcel_plomb_status'] == 'Принят'])
        quon_not_done = quont_all_plombs - quont_plomb_done
        object_name = parcel_plomb_numb
        comment = f'Приемка по местам: Пломба принята (Партия: {party_numb})'
        insert_user_action(object_name, comment)
        print(df)
        return index, quont_all_plombs, quont_plomb_done, quon_not_done, vector, df, parcel_plomb_numb
