import sqlite3
import sqlite3 as sl
import traceback

import numpy as np
import pandas as pd
import winsound
from flask import flash
from flask import request, render_template, Blueprint
from flask_jwt_extended import jwt_required

from SVH_BAZA_modules.services import (insert_user_action, map_eng_to_rus, addition_folder,
                                       logger, get_user_name)
from . services import style


bp_pallet = Blueprint('pallet', __name__, url_prefix='/pallet')

parcel_plomb_numb_np = None
df_plombs_np = pd.DataFrame()
vector = None


@bp_pallet.route('/create_new_pallet', methods=['POST', 'GET'])
def create_pallet():
    global df_plombs_html
    global df_plombs_np
    global parcel_plomb_numb_np
    global vector
    modal = 0
    try:
        parcel_plomb_numb_np = request.form['parcel_plomb_numb_np']
    except:
        pass
    try:
        con = sl.connect('BAZA.db')
        with con:
            df_all_pallets = pd.read_sql(f"SELECT DISTINCT pallet FROM baza", con).drop_duplicates(subset='pallet',
                                                                                                   keep='first')
            try:
                df_all_pallets['pallet'] = df_all_pallets['pallet'].fillna(0).astype(int)
                df_all_pallets = df_all_pallets.sort_values(by='pallet')
                last_pall_numb = int(df_all_pallets.values[-1].tolist()[0])
                i = last_pall_numb + 1
            except:
                last_pall_numb = df_all_pallets.values[0].tolist()[0]
                i = 1
            logger.warning(last_pall_numb)
            df_plomb = pd.read_sql(f"SELECT * FROM baza where parcel_plomb_numb = '{parcel_plomb_numb_np}'", con)
            try:
                first_parcel = df_plomb['parcel_numb'].values[0]
                print(first_parcel)
            except:
                pass
            if df_plomb['custom_status_short'].astype(str).str.contains('ИЗЪЯТИЕ').any():
                flash(f'ВНИМАНИЕ, пломба с неотвязанными посылками на изъятие!', category='error')
                winsound.PlaySound('Snd\Snd_CancelIssue.wav', winsound.SND_FILENAME)
            df_plomb = df_plomb.drop_duplicates(subset=['parcel_plomb_numb'], keep='first')
            df_plomb['Тип'] = 'Пломба'
            if df_plomb.empty and parcel_plomb_numb_np is not None:
                df_plomb = pd.read_sql(f"SELECT * FROM baza where parcel_numb = '{parcel_plomb_numb_np}'", con)
                df_plomb['Тип'] = "Посылка-место"
                logger.warning(df_plomb)
                if df_plomb.empty and parcel_plomb_numb_np is not None:
                    flash(f'Пломба не найдена!', category='error')
                    winsound.PlaySound('Snd\Snd_NoPlomb.wav', winsound.SND_FILENAME)
                else:
                    pass
            else:
                pass

            try:
                df_plombs_np = df_plombs_np.append(df_plomb).drop_duplicates(subset=['parcel_plomb_numb'], keep='last')
                quont_plombs = len(df_plombs_np)
                df_plombs_np.index = df_plombs_np.index + 1  # shifting index
                df_plombs_np.sort_index(inplace=True)
                df_plombs_np['№1'] = np.arange(len(df_plombs_np))[::-1] + 1
                df_plombs_html = df_plombs_np.rename(columns={
                    '№1': '№',
                    'parcel_plomb_numb': 'Пломба',
                    'parcel_numb': 'Трек',
                    'pallet': '№ Паллет',
                    'zone': 'Зона',
                    'party_numb': 'Партия'
                })
                df_plombs_html = df_plombs_html[['№', 'Пломба', '№ Паллет',
                                                 'Зона', 'Партия', 'Трек', 'Тип']]
                df_plombs_html.fillna("", inplace=True)
                df_plombs_html = df_plombs_html.reset_index()
                df_plombs_html = df_plombs_html.drop('index', axis=1)

                object_name = parcel_plomb_numb_np
                comment = f'Сформировать новый паллет: пломба отмечена для добавления на паллет №{i}'
                insert_user_action(object_name, comment)
                winsound.PlaySound('Snd\zvuk-vezeniya.wav', winsound.SND_FILENAME)
            except:
                pass
        try:
            con = sl.connect('VECTORS.db')
            with con:
                try:
                    try:
                        vector1 = con.execute(
                            f"SELECT vector FROM VECTORS where parcel_plomb_numb = '{parcel_plomb_numb_np}'").fetchone()[
                            0]
                        flash(f'{vector1}!', category='success')
                        print(vector1)
                    except:
                        try:
                            vector1 = con.execute(
                                f"SELECT vector FROM VECTORS where parcel_numb = '{first_parcel}'").fetchone()[
                                0]
                            flash(f'{vector1}!', category='success')
                            print(vector1)
                        except Exception as e:
                            vector1 = None
                            print(e)

                    if vector != vector1 and vector is not None:
                        flash(f'Направления не совпадают!! Было ({vector})!', category='error')
                        winsound.PlaySound('Snd\Snd_CancelIssue.wav', winsound.SND_FILENAME)
                        modal = 1
                    vector = vector1
                except Exception as e:
                    print(e)

        except Exception as e:
            flash(f'{e}!', category='error')


    except Exception as e:
        flash(f'Пломба не найдена! {e}', category='error')
        winsound.PlaySound('Snd\Snd_NoPlomb.wav', winsound.SND_FILENAME)
        # return render_template('parcel_info_new_place.html')
        return {'message': str(e)}, 400

    return render_template('New_pallet.html', tables=[style + df_plombs_html.to_html(classes='mystyle', index=False,
                                                                                     float_format='{:2,.2f}'.format)],
                           titles=['na', '\n\nМешки\места:'],
                           parcel_plomb_numb_np=parcel_plomb_numb_np, i=i, quont_plombs=quont_plombs,
                           modal=modal, vector=vector)


df_plombs_html = pd.DataFrame()


@bp_pallet.route('/create_new_pallet_sql', methods=['POST', 'GET'])
def create_pallet_sql():
    user_name, user_id = get_user_name()
    audiofile = 'None'
    modal = 0
    first_parcel = None
    vector = None
    try:
        parcel_plomb_numb = request.form['parcel_plomb_numb']
    except:
        parcel_plomb_numb = None
    try:
        con_vect = sl.connect('VECTORS.db')
        con = sl.connect('BAZA.db')
        with con:
            df_all_pallets = pd.read_sql(f"SELECT DISTINCT pallet FROM baza", con).drop_duplicates(subset='pallet',
                                                                                           keep='first')
            df_plomb = pd.read_sql(f"SELECT * FROM baza where parcel_plomb_numb = '{parcel_plomb_numb}'", con)
            try:
                first_parcel = df_plomb['parcel_numb'].values[0]
                print(first_parcel)
            except:
                pass
        with con_vect:
            try:
                vector = con_vect.execute(
                    f"SELECT vector FROM vectors where parcel_plomb_numb = '{parcel_plomb_numb}'").fetchone()[
                    0]
                flash(f'{vector}!', category='success')
                print(vector)
            except:
                try:
                    vector = con_vect.execute(
                        f"SELECT vector FROM vectors where parcel_numb = '{first_parcel}'").fetchone()[
                        0]
                    flash(f'{vector}!', category='success')
                    print(vector)
                except Exception as e:
                    vector = None
                    print(e)
            try:
                df_all_pallets['pallet'] = df_all_pallets['pallet'].fillna(0).astype(int)
                df_all_pallets = df_all_pallets.sort_values(by='pallet')
                last_pall_numb = int(df_all_pallets.values[-1].tolist()[0])
                i = last_pall_numb + 1
            except:
                last_pall_numb = df_all_pallets.values[0].tolist()[0]
                i = 1

            if df_plomb['custom_status_short'].astype(str).str.contains('ИЗЪЯТИЕ').any():
                flash(f'ВНИМАНИЕ, пломба с изъятием!', category='error')
                winsound.PlaySound('Snd\Snd_CancelIssue.wav', winsound.SND_FILENAME)
            df_plomb = df_plomb.drop_duplicates(subset=['parcel_plomb_numb'], keep='first')
            if df_plomb.empty and parcel_plomb_numb is not None:
                flash(f'Пломба не найдена!', category='error')
                winsound.PlaySound('Snd\Snd_NoPlomb.wav', winsound.SND_FILENAME)
            if not df_plomb.empty:
                try:
                    with con:
                        con.execute(f"INSERT OR REPLACE INTO add_to_pallet (parcel_plomb_numb, user_id, vector) "
                                    f"VALUES('{parcel_plomb_numb}', '{user_id}', '{vector}')")
                except Exception as e:
                    flash(f'Пломба не найдена! {e}', category='error')
                    audiofile = 'Snd\Snd_NoPlomb.wav'
                    print(str(traceback.format_exc()))
        with con:
            df_plombs_np = pd.read_sql(f"SELECT * FROM add_to_pallet where user_id = '{user_id}'",
                                       con).fillna('').sort_values(by='ID', ascending=False)
        quont_plombs = len(df_plombs_np)
        df_plombs_np['№1'] = np.arange(len(df_plombs_np))[::-1] + 1
        df_plombs_html = df_plombs_np.rename(columns={
            '№1': '№',
            'parcel_plomb_numb': 'Пломба',
            'vector': 'Направление',
            'user_id': 'user_id'
        })
        print(df_plombs_html)
        df_plombs_html = df_plombs_html[['№', 'Пломба', 'Направление', 'user_id']]
        df_plombs_html.fillna("", inplace=True)

        df_vectors = df_plombs_html[['Пломба', 'Направление']].drop_duplicates(subset='Направление')
        if len(df_vectors) > 1:
            flash(f'Пломбы с разными направлениями!', category='error')
        object_name = parcel_plomb_numb
        comment = f'Сформировать новый паллет: пломба отмечена для добавления на паллет №{i}'
        insert_user_action(object_name, comment)
        audiofile = 'Snd\zvuk-vezeniya.wav'
        return render_template('New_pallet.html', tables=[style + df_plombs_html.to_html(classes='mystyle', index=False,
                                                                                         float_format='{:2,.2f}'.format)],
                               titles=['na', '\n\nМешки\места:'],
                               parcel_plomb_numb=parcel_plomb_numb, i=i, quont_plombs=quont_plombs,
                               modal=modal, vector=vector, audiofile=audiofile)
    except Exception as e:
        #winsound.PlaySound('Snd\Snd_NoPlomb.wav', winsound.SND_FILENAME)
        audiofile = 'Snd\Snd_NoPlomb.wav'
        print(str(traceback.format_exc()))
        return render_template('New_pallet.html', audiofile=audiofile)



@bp_pallet.route('/delete_last_plomb_sql', methods=['POST', 'GET'])
def delete_last_plomb_sql():
    user_name, user_id = get_user_name()
    con = sqlite3.connect('BAZA.db')
    with con:
        df_last_plomb = pd.read_sql(f"SELECT * FROM add_to_pallet WHERE user_id = '{user_id}' AND id = (SELECT MAX(id) FROM add_to_pallet)", con)
        deleted_plomb_numb = df_last_plomb['parcel_plomb_numb'].values[0]
        con.execute(f"DELETE FROM add_to_pallet WHERE parcel_plomb_numb = '{deleted_plomb_numb}'")
        df_plombs_np = pd.read_sql(f"SELECT * FROM add_to_pallet where user_id = '{user_id}'",
                                   con).fillna('').sort_values(by='ID', ascending=False)

    df_plombs_np['№1'] = np.arange(len(df_plombs_np))[::-1] + 1
    df_plombs_html = df_plombs_np.rename(columns={
        '№1': '№',
        'parcel_plomb_numb': 'Пломба',
        'vector': 'Направление',
        'user_id': 'user_id'
    })
    print(df_plombs_html)
    df_plombs_html = df_plombs_html[['№', 'Пломба', 'Направление', 'user_id']]
    df_plombs_html.fillna("", inplace=True)
    object_name = deleted_plomb_numb
    quont_plombs = len(df_plombs_np)
    comment = f'Удалена пломба из создания паллета'
    insert_user_action(object_name, comment)
    flash(f'{deleted_plomb_numb} Удалена из списка добавления', category='success')
    return render_template('New_pallet.html',
                           tables=[style + df_plombs_html.to_html(classes='mystyle',
                                                                                     index=False,
                                                                                     float_format='{:2,.2f}'.format)],
                           titles=['na', '\n\nМешки\места:'],
                           parcel_plomb_numb_np=parcel_plomb_numb_np, quont_plombs=quont_plombs)


@bp_pallet.route('/delete_last_plomb', methods=['POST', 'GET'])
def delete_last_plomb():
    global parcel_plomb_numb_np
    global df_plombs_np
    global df_plombs_html
    logger.warning(df_plombs_html)
    deleted_plomb_numb = df_plombs_np['parcel_plomb_numb'].iloc[0]
    df_plombs_html = df_plombs_html[1:]
    df_plombs_np = df_plombs_np[1:]
    object_name = deleted_plomb_numb
    comment = f'Удалена пломба из создания паллета'
    insert_user_action(object_name, comment)
    return render_template('New_pallet.html',
                           tables=[style + df_plombs_html.to_html(classes='mystyle',
                                                                                     index=False,
                                                                                     float_format='{:2,.2f}'.format)],
                           titles=['na', '\n\nМешки\места:'],
                           parcel_plomb_numb_np=parcel_plomb_numb_np)



@bp_pallet.route('/delete_last_plomb_addpallet', methods=['POST', 'GET'])
def delete_last_plomb_addpallet():
    global parcel_plomb_numb_np
    global df_plombs_np
    global df_plombs_html
    logger.warning(df_plombs_html)
    df_plombs_html = df_plombs_html[:-1]
    df_plombs_np = df_plombs_np[:-1]
    object_name = parcel_plomb_numb_np
    flash(f'{parcel_plomb_numb_np} Удалена из добавления на паллет', category='success')
    comment = f'Удалена пломба из добавления на паллет'
    insert_user_action(object_name, comment)
    return render_template('add_to_pallet.html',
                           tables=[style + df_plombs_html.to_html(classes='mystyle',
                                                                  index=False,
                                                                                        float_format='{:2,.2f}'.format)],
                           titles=['na', '\n\nМешки\места:'],
                           parcel_plomb_numb_np=parcel_plomb_numb_np)


@bp_pallet.route('/search/clean_working_place_pallet_sql', methods=['POST'])
def clean_working_place_pallet_sql():
    user_name, user_id = get_user_name()
    con = sqlite3.connect('BAZA.db')
    with con:
        df_plomb = pd.read_sql(f"SELECT * FROM add_to_pallet WHERE user_id = '{user_id}'", con)
        list_of_plombs = df_plomb['parcel_plomb_numb'].to_list()
        qt_plomb = len(list_of_plombs)
        con.execute(f"DELETE FROM add_to_pallet WHERE user_id = '{user_id}'")
    flash(f'Таблица очищена, удалено {qt_plomb} пломб: {str(list_of_plombs)}', category='success')
    object_name = None
    comment = f'Сформировать новый паллет: Таблица очищена, удалено {qt_plomb} пломб: {str(list_of_plombs)}'
    insert_user_action(object_name, comment)
    return render_template('New_pallet.html')


@bp_pallet.route('/search/clean_working_place_pallet', methods=['POST'])
def clean_working_place_pallet():
    df_plombs_html = pd.DataFrame()
    df_plombs_np = pd.DataFrame()
    parcel_plomb_numb_np = None
    object_name = None
    comment = f'Сформировать новый паллет: Таблица очищена'
    insert_user_action(object_name, comment)
    return render_template('New_pallet.html')



@bp_pallet.route('/search/clean_working_place_clean_working_place_addpallet', methods=['POST'])
def clean_working_place_addpallet():
    global parcel_plomb_numb_np
    global df_plombs_np
    global df_plombs_html
    df_plombs_html = pd.DataFrame()
    df_plombs_np = pd.DataFrame()
    object_name = None
    comment = f'Добавить на паллет: Таблица очищена'
    insert_user_action(object_name, comment)
    return render_template('add_to_pallet.html')


@bp_pallet.route('/insert_new_pallet', methods=['POST', 'GET'])
def insert_pallet():
    global df_plombs_np
    global parcel_plomb_numb_np
    pallet_new = request.form['pallet_new']
    con = sl.connect('BAZA.db')
    if pallet_new is None or pallet_new == '':
        with con:
            df_all_pallets = pd.read_sql(f"SELECT pallet FROM baza", con).drop_duplicates(subset='pallet', keep='first')
            try:
                print(df_all_pallets['pallet'])
                df_all_pallets['pallet'] = df_all_pallets['pallet'].fillna(0).astype(int)
                df_all_pallets = df_all_pallets.sort_values(by='pallet')
                logger.warning(df_all_pallets)
                last_pall_numb = int(df_all_pallets.values[-1].tolist()[0])
                pallet_new = last_pall_numb + 1
            except Exception as e:
                print(e)
    else:
        pass
    # открываем базу
    con = sl.connect('BAZA.db')
    with con:
        df_plombs_np_typeplomb = df_plombs_np.loc[df_plombs_np['Тип'] == 'Пломба']
        logger.warning(df_plombs_np_typeplomb)
        df_plombs_np_typeparcel = df_plombs_np.loc[df_plombs_np['Тип'] == 'Посылка-место']
        logger.warning(df_plombs_np_typeparcel)
        for parcel_plomb_numb in df_plombs_np_typeplomb['parcel_plomb_numb']:
            con.execute(
                f"Update baza set pallet = '{pallet_new}' where parcel_plomb_numb = '{parcel_plomb_numb}'")
        for parcel_numb in df_plombs_np_typeparcel['parcel_numb']:
            con.execute(
                f"Update baza set pallet = '{pallet_new}' where parcel_numb = '{parcel_numb}'")
        df_new_pallet = pd.read_sql(f"Select * from baza where pallet = '{pallet_new}'", con)
        writer = pd.ExcelWriter(f'{addition_folder}Паллет {pallet_new}.xlsx', engine='xlsxwriter')
        df_new_pallet.to_excel(writer, sheet_name='Sheet1', index=False)
        for column in df_new_pallet:
            column_width = max(df_new_pallet[column].astype(str).map(len).max(), len(column))
            col_idx = df_new_pallet.columns.get_loc(column)
            writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
            writer.sheets['Sheet1'].set_column(0, 3, 10)
            writer.sheets['Sheet1'].set_column(1, 3, 20)
            writer.sheets['Sheet1'].set_column(2, 3, 20)
            writer.sheets['Sheet1'].set_column(3, 3, 20)
            writer.sheets['Sheet1'].set_column(4, 3, 30)
            writer.sheets['Sheet1'].set_column(5, 3, 20)

        writer.save()
        con.commit()
        df_plombs_np = pd.DataFrame()
        parcel_plomb_numb_np = None
        object_name = pallet_new
        comment = f'Сформировать новый паллет: Паллет сформирован'
        insert_user_action(object_name, comment)
        flash(f'Паллет {pallet_new} сформирован!', category='success')
        winsound.PlaySound('Snd\Pallet_made.wav', winsound.SND_FILENAME)
    return render_template('New_pallet.html')


@bp_pallet.route('/insert_pallet_sql', methods=['POST', 'GET'])
def insert_pallet_sql():
    user_name, user_id = get_user_name()
    #pallet_new = request.form['pallet_new']
    con = sl.connect('BAZA.db')
    with con:
        df_all_pallets = pd.read_sql(f"SELECT pallet FROM baza", con).drop_duplicates(subset='pallet', keep='first')
        try:
            print(df_all_pallets['pallet'])
            df_all_pallets['pallet'] = df_all_pallets['pallet'].fillna(0).astype(int)
            df_all_pallets = df_all_pallets.sort_values(by='pallet')
            logger.warning(df_all_pallets)
            last_pall_numb = int(df_all_pallets.values[-1].tolist()[0])
            pallet_new = last_pall_numb + 1
        except Exception as e:
            print(e)
    # открываем базу
    with con:
        df_plomb = pd.read_sql(f"SELECT * FROM add_to_pallet WHERE user_id = '{user_id}'", con)
        for parcel_plomb_numb in df_plomb['parcel_plomb_numb']:
            con.execute(
                f"Update baza set pallet = '{pallet_new}' where parcel_plomb_numb = '{parcel_plomb_numb}'")
        con.execute(
            f"DELETE FROM add_to_pallet WHERE user_id = '{user_id}'")
        df_new_pallet = pd.read_sql(f"Select * from baza where pallet = '{pallet_new}'", con)
        writer = pd.ExcelWriter(f'{addition_folder}Паллет {pallet_new}.xlsx', engine='xlsxwriter')
        df_new_pallet.to_excel(writer, sheet_name='Sheet1', index=False)
        for column in df_new_pallet:
            column_width = max(df_new_pallet[column].astype(str).map(len).max(), len(column))
            col_idx = df_new_pallet.columns.get_loc(column)
            writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
            writer.sheets['Sheet1'].set_column(0, 3, 10)
            writer.sheets['Sheet1'].set_column(1, 3, 20)
            writer.sheets['Sheet1'].set_column(2, 3, 20)
            writer.sheets['Sheet1'].set_column(3, 3, 20)
            writer.sheets['Sheet1'].set_column(4, 3, 30)
            writer.sheets['Sheet1'].set_column(5, 3, 20)

        writer.save()
        con.commit()

    object_name = pallet_new
    comment = f'Сформировать новый паллет: Паллет сформирован'
    insert_user_action(object_name, comment)
    flash(f'Паллет {pallet_new} сформирован!', category='success')
    return render_template('New_pallet.html')


@bp_pallet.route('/add_to_pallet', methods=['POST', 'GET'])
def add_to_pallet():
    global df_plombs_html
    global df_plombs_np
    global parcel_plomb_numb_np
    try:
        parcel_plomb_numb_np = request.form['parcel_plomb_numb_np']
    except:
        pass
    try:
        con = sl.connect('BAZA.db')
        with con:
            # to show all pallets from system for choice
            df_all_pallets = pd.read_sql(f"SELECT pallet FROM baza", con).drop_duplicates(subset='pallet', keep='first')
            df_all_pallets['pallet'] = df_all_pallets['pallet'].fillna(value=np.nan).fillna(0).astype(int)
            df_all_pallets = df_all_pallets.sort_values(by='pallet', na_position='last', ascending=False)
            df_all_pallets = df_all_pallets['pallet'].to_list()
            logger.warning(df_all_pallets)
            # select row with current plomb (parcel) number
            df_plomb = pd.read_sql(f"SELECT * FROM baza where parcel_plomb_numb = '{parcel_plomb_numb_np}'", con)
            df_plomb['Тип'] = 'Пломба'
            # if that is parcel
            if df_plomb.empty and parcel_plomb_numb_np != None:
                df_plomb = pd.read_sql(f"SELECT * FROM baza where parcel_numb = '{parcel_plomb_numb_np}'", con)
                df_plomb['Тип'] = "Посылка-место"
                logger.warning(df_plomb)
                if df_plomb.empty and parcel_plomb_numb_np != None:
                    flash(f'Пломба не найдена!', category='error')
                    winsound.PlaySound('Snd\Snd_NoPlomb.wav', winsound.SND_FILENAME)
            else:
                pass
            try:
                df_plombs_np = df_plombs_np.append(df_plomb).drop_duplicates(subset=['parcel_plomb_numb'], keep='first')
                df_plombs_np['№1'] = np.arange(len(df_plombs_np))[::+1] + 1
                df_plombs_html = df_plombs_np.rename(columns={
                    '№1': '№',
                    'parcel_plomb_numb': 'Пломба',
                    'parcel_numb': 'Трек',
                    'pallet': '№ Паллет',
                    'zone': 'Зона',
                    'party_numb': 'Партия'
                })
                df_plombs_html = df_plombs_html[['№', 'Пломба', '№ Паллет',
                                                 'Зона', 'Партия', 'Трек', 'Тип']]
                df_plombs_html.fillna("", inplace=True)
                df_plombs_html = df_plombs_html.reset_index()
                df_plombs_html = df_plombs_html.drop('index', axis=1)
                object_name = parcel_plomb_numb_np
                comment = f'Добавить на паллет: Пломба отмечена для добавления'
                insert_user_action(object_name, comment)
            except:
                pass

    except Exception as e:
        flash(f'Пломба не найдена!', category='error')
        winsound.PlaySound('Snd\Snd_NoPlomb.wav', winsound.SND_FILENAME)
        # return render_template('parcel_info_new_place.html')
        return {'message': str(e)}, 400
        pass
    return render_template('add_to_pallet.html', tables=[style + df_plombs_html.to_html(classes='mystyle', index=False,
                                                                                        float_format='{:2,.2f}'.format)],
                           titles=['na', '\n\nМешки\места:'],
                           parcel_plomb_numb_np=parcel_plomb_numb_np, df_all_pallets=df_all_pallets)


@bp_pallet.route('/add_to_pallet_button', methods=['POST', 'GET'])
def add_to_pallet_button():
    global df_plombs_np
    global parcel_plomb_numb_np
    pallet_new = request.form['pallet_new']
    # открываем базу
    con = sl.connect('BAZA.db')
    with con:
        df_new_pallet = pd.read_sql(f"Select * from baza where pallet = '{pallet_new}'", con)
        logger.warning(df_new_pallet)
        df_plombs_np_typeplomb = df_plombs_np.loc[df_plombs_np['Тип'] == 'Пломба']
        logger.warning(df_plombs_np_typeplomb)
        df_plombs_np_typeparcel = df_plombs_np.loc[df_plombs_np['Тип'] == 'Посылка-место']
        logger.warning(df_plombs_np_typeparcel)
        if not df_new_pallet.empty:
            for parcel_plomb_numb in df_plombs_np_typeplomb['parcel_plomb_numb']:
                # plomb_toreplace = done_parcels_np.loc[done_parcels_np['Трек-номер'] == parcel_numb]['Пломба'].values[0]
                # logger.warning(plomb_toreplace)
                con.execute(
                    f"Update baza set pallet = '{pallet_new}' where parcel_plomb_numb = '{parcel_plomb_numb}'")
            for parcel_numb in df_plombs_np_typeparcel['parcel_numb']:
                con.execute(
                    f"Update baza set pallet = '{pallet_new}' where parcel_numb = '{parcel_numb}'")
            df_new_pallet = pd.read_sql(f"Select * from baza where pallet = '{pallet_new}'", con)
            writer = pd.ExcelWriter(f'{addition_folder}Паллет {pallet_new}.xlsx', engine='xlsxwriter')
            df_new_pallet.to_excel(writer, sheet_name='Sheet1', index=False)
            for column in df_new_pallet:
                column_width = max(df_new_pallet[column].astype(str).map(len).max(), len(column))
                col_idx = df_new_pallet.columns.get_loc(column)
                writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
                writer.sheets['Sheet1'].set_column(0, 3, 10)
                writer.sheets['Sheet1'].set_column(1, 3, 20)
                writer.sheets['Sheet1'].set_column(2, 3, 20)
                writer.sheets['Sheet1'].set_column(3, 3, 20)
                writer.sheets['Sheet1'].set_column(4, 3, 30)
                writer.sheets['Sheet1'].set_column(5, 3, 20)

            writer.save()
            con.commit()
            df_plombs_np = pd.DataFrame()
            parcel_plomb_numb_np = None
            object_name = pallet_new
            comment = f'Добавить на паллет: Пломбы добавленны'
            insert_user_action(object_name, comment)
            flash(f'Добавлено!', category='success')
            winsound.PlaySound('Snd\dd_.wav', winsound.SND_FILENAME)
        else:
            flash(f'Паллет не найден!', category='error')
            winsound.PlaySound('Snd\Pallet_not_found.wav', winsound.SND_FILENAME)
    return render_template('add_to_pallet.html')


parc_quont_pallet_info = 0
plomb_quont_pallet_info = 0
pallet = 0

@bp_pallet.route('/pallet_info', methods=['POST', 'GET'])
def pallet_info():
    global parc_quont_pallet_info
    global plomb_quont_pallet_info
    global pallet
    df_refuses = []
    numb = 0
    try:
        numb = request.form['numb']
        logger.warning(numb)
    except:
        pass
    con = sl.connect('BAZA.db')
    if numb == 0:
        df = pd.DataFrame()
    else:
        try:
            with con:
                df = pd.read_sql(f"Select * from baza where pallet = '{numb}'", con)
                df = df.rename(columns=map_eng_to_rus)
                df['№'] = np.arange(len(df))[::+1] + 1
                df = df[
                    ['№', 'Паллет', 'Пломба', 'Трек-номер', 'Статус ВХ', 'Статус ТО (кратк)', 'Статус ТО', 'Партия']]
                if df.empty:
                    df = pd.read_sql(f"Select * from baza where parcel_plomb_numb = '{numb}'", con)
                    try:
                        pallet = df['pallet'].values[0]
                        df = pd.read_sql(f"Select * from baza where pallet = '{pallet}'", con)
                        df = df.rename(columns=map_eng_to_rus)
                        df['№'] = np.arange(len(df))[::+1] + 1
                        df = df[['№', 'Паллет', 'Пломба', 'Трек-номер', 'Статус ВХ', 'Статус ТО (кратк)', 'Статус ТО',
                                 'Партия']]
                    except:
                        pass
                    if df.empty:
                        df = pd.read_sql(f"Select * from baza where parcel_numb = '{numb}'", con)
                        try:
                            pallet = df['pallet'].values[0]
                            df = pd.read_sql(f"Select * from baza where pallet = '{pallet}'", con)
                            df = df.rename(columns=map_eng_to_rus)
                            df['№'] = np.arange(len(df))[::+1] + 1
                            df = df[
                                ['№', 'Паллет', 'Пломба', 'Трек-номер', 'Статус ВХ', 'Статус ТО (кратк)', 'Статус ТО',
                                 'Партия']]
                        except:
                            pass
                df_refuses = df.loc[df['Статус ТО (кратк)'] == 'ИЗЪЯТИЕ']
                df_refuses = df_refuses['Трек-номер'].to_list()
                logger.warning(df_refuses)
                parc_quont_pallet_info = len(df)
                plomb_quont_pallet_info = len(df.drop_duplicates(subset='Пломба', keep='first'))
                pallet = df['Паллет'].values[0]

                def highlight_RED(df):
                    return ['background-color: #FF0000' if 'ИЗЪЯТИЕ' in str(i) else '' for i in df]

                df = df.style.apply(highlight_RED).hide_index()
        except KeyError as ke:
            flash(f'Паллет не найден! {ke}', category='error')
    return render_template('pallet_info.html', tables=[style + df.to_html(classes='mystyle', index=False,
                                                                          float_format='{:2,.2f}'.format)],
                           titles=['Информация'], df_refuses=df_refuses,
                           pallet=pallet, numb=numb, parc_quont_pallet_info=parc_quont_pallet_info,
                           plomb_quont_pallet_info=plomb_quont_pallet_info)


@bp_pallet.route('/pallet_info_callback_refuses', methods=['POST', 'GET'])
def pallet_info_callback_refuses():
    global pallet
    con = sl.connect('BAZA.db')
    with con:
        df_plombs = pd.read_sql(
            f"Select parcel_plomb_numb from baza where pallet = '{pallet}' AND custom_status_short = 'ИЗЪЯТИЕ'", con)
        df_plombs = df_plombs.drop_duplicates()
        df_plombs.to_sql('plombs_with_pullings', con, if_exists='append', index=False)
        con.execute(
            f"Update baza set pallet = '0', parcel_plomb_numb = '' where pallet = '{pallet}' AND custom_status_short = 'ИЗЪЯТИЕ'")
        print('updated')
        object_name = pallet
        comment = f'Отвязанны посылки с изъятием от паллета'
        insert_user_action(object_name, comment)
    flash(f'Посылки со статусом ИЗЪЯТИЕ успешно отвязаны от паллета {pallet}! Убедитесь, что их реально изъяли!',
          category='success')
    return render_template('pallet_info.html', pallet=pallet)
