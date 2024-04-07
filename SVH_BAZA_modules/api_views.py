import json
import pandas as pd
import sqlite3 as sl
import pytz
import winsound
from flask import Blueprint, abort, Response
import requests
from apscheduler.schedulers.background import BackgroundScheduler
from flask import jsonify, request, render_template, flash
import datetime
from SVH_BAZA_modules.services import (insert_user_action, addition_folder, logger)

bp_api = Blueprint('api', __name__, url_prefix='/api')


def insert_decision_API(parcel_numb, custom_status, custom_status_short, refuse_reason, decision_date):
    conn = sl.connect('BAZA.db')
    cur = conn.cursor()
    statement = ("INSERT INTO events2 (parcel_numb, custom_status, "
                 "custom_status_short, refuse_reason, decision_date) VALUES (?, ?, ?, ?, ?)")
    cur.execute(statement, [parcel_numb, custom_status, custom_status_short,
                            refuse_reason, decision_date])
    conn.commit()
    conn.close()
    return True


@bp_api.route('/api/add_decision', methods=['POST'])
def add_decision_API():
    event_details = request.get_json()
    parcel_numb = event_details["parcel_numb"]
    custom_status = event_details["Event"]
    if 'Выпуск' in str(custom_status):
        custom_status_short = 'ВЫПУСК'
    else:
        custom_status_short = 'ИЗЪЯТИЕ'
    refuse_reason = event_details["Event_comment"]
    decision_date = datetime.datetime.strptime(event_details["Event_date"], "%Y-%m-%d %H:%M:%S").replace(
        tzinfo=pytz.UTC).astimezone(pytz.timezone("Europe/London"))
    result = insert_decision_API(parcel_numb, custom_status, custom_status_short, refuse_reason, decision_date)
    return jsonify(result)


def insert_event_API_test(df, event_date):
    parcel_list = df['parcel_numb'].to_list()
    logger.warning(parcel_list)
    logger.warning(event_date)
    for parcel in parcel_list:
        body = {"parcel_numb": parcel, "Event": "Отгружен с таможенного склада",
                "Event_comment": "Уссурийск", "Event_date": event_date}
        headers = {'accept': 'application/json'}
        response = requests.post('http://164.132.182.145:5000/api/add/new_event', json=body,
                                 headers={'accept': 'application/json'})
        try:
            return response.json()
        except ValueError:
            pass


@bp_api.route('/webhook', methods=['POST'])
def get_webhook():
    if request.method == 'POST':
        print("received data: ", request.json)
        return 'success', 200
    else:
        abort(400)


@bp_api.route('/api/del_plomb_TSD', methods=['POST'])
def api_del_plomb_TSD():
    parcel_details = request.get_json()
    parcel_plomb_numb = parcel_details['parcel_plomb_numb']
    con = sl.connect('BAZA.db')
    print(parcel_plomb_numb)
    try:
        with con:
            con.execute(
                f"Update baza set VH_status = 'На ВХ', parcel_plomb_numb = ''  "
                f"where parcel_plomb_numb = '{parcel_plomb_numb}' "
                f"and custom_status_short = 'ИЗЪЯТИЕ'")
            print('ok')

    except Exception as e:
        print(e)
        return {'message': str(e)}, 400
    return jsonify('True')


@bp_api.route('/api/get_plomb_info_API_TSD', methods=['POST'])
def get_plomb_info_API_TSD():
    plomb_details = request.get_json()
    parcel_plomb_numb = plomb_details['parcel_plomb_numb']
    print(parcel_plomb_numb)
    con = sl.connect('BAZA.db')
    try:
        with con:
            query_plomb_status = f"SELECT parcel_numb, custom_status_short from baza where parcel_plomb_numb = '{parcel_plomb_numb}'"
            df_plomb_status = pd.read_sql(query_plomb_status, con)
            print(df_plomb_status)
    except Exception as e:
        print(e)
        return {'message': str(e)}, 400

    return Response(df_plomb_status.to_json(orient="records", indent=2), mimetype='application/json')


@bp_api.route('/api/create_pallet_API_TSD', methods=['POST'])
def create_pallet_API_TSD():
    plomb_details = request.get_json()
    print(plomb_details)
    df_plombs = pd.json_normalize(json.loads(plomb_details))
    print(df_plombs)
    try:
        con = sl.connect('BAZA.db')
        with con:
            df_all_pallets = pd.read_sql(f"SELECT DISTINCT pallet FROM baza", con).drop_duplicates(subset='pallet',
                                                                                                   keep='first')
            try:
                df_all_pallets['pallet'] = df_all_pallets['pallet'].fillna(0).astype(int)
                df_all_pallets = df_all_pallets.sort_values(by='pallet')
                last_pall_numb = int(df_all_pallets.values[-1].tolist()[0])
                pallet_new = last_pall_numb + 1
            except:
                pallet_new = 1
            for parcel_plomb_numb in df_plombs['parcel_plomb_numb']:
                print(parcel_plomb_numb)
                con.execute(
                    f"Update baza set pallet = '{pallet_new}' where parcel_plomb_numb = '{parcel_plomb_numb}'")
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
            flash(f'Паллет сформирован!', category='success')
            winsound.PlaySound('Snd\Pallet_made.wav', winsound.SND_FILENAME)
        result = jsonify({"pallet": pallet_new})
    except Exception as e:
        result = jsonify({"Error": e})
    return result


@bp_api.route('/api/update_decisions/<party_numb>', methods=['POST', 'GET'])
def API_update_decisions(party_numb):
    con = sl.connect('BAZA.db')
    with con:
        df_request_decisions = pd.read_sql(f"SELECT parcel_numb FROM baza WHERE party_numb = '{party_numb}'", con)
    body = df_request_decisions.to_json(orient="records", indent=2)
    # body = {'parcel_numb': 'CEL7000012753CD'}
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
                    custom_status_short = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb][
                        'custom_status_short'].values[0]
                    custom_status = \
                        df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb][
                            'custom_status'].values[
                            0]
                    decision_date = \
                        df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb][
                            'decision_date'].values[
                            0]
                    refuse_reason = \
                        df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb][
                            'refuse_reason'].values[
                            0]

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
            flash(f'Решения по партии {party_numb} загружены')
            object_name = party_numb
            comment = 'Решения по партии загружены вручную'
            insert_user_action(object_name, comment)
    except:
        print(response)
        flash(f'Ошибка загрузки решений {response}')
    return render_template('index.html')


def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def server_request_events():
    print("start updating")
    con = sl.connect('BAZA.db')
    len_id = con.execute('SELECT max(id) from baza').fetchone()[0]
    id_for_job = len_id - 700000
    df = pd.read_sql(f"Select parcel_numb from baza "
                     f"where ID > {id_for_job} "
                     f"AND custom_status_short = 'ИЗЪЯТИЕ' ", con).drop_duplicates(subset='parcel_numb')


    list_chanks = list(chunks(df, 100))
    # print(list_chanks)
    i = 0
    for chank in list_chanks:
        i += 1
        with con:
            body = chank.to_json(orient="records", indent=2)
            # body = {'parcel_numb': 'CEL7000012753CD'}
            headers = {'accept': 'application/json'}
            response = requests.post('http://164.132.182.145:5001/api/get_decisions',
                                     json=body)  # http://127.0.0.1:5000  # 'http://164.132.182.145:5001/api/get_decisions'
            try:
                json_decisions = response.json()
                df_loaded_decisions = pd.DataFrame.from_records(json_decisions)
                with con:
                    df_to_append = pd.DataFrame()

                    for parcel_numb in df_loaded_decisions['parcel_numb']:

                        row_isalready_in = pd.read_sql(f"Select * from baza where parcel_numb = '{parcel_numb}'", con)
                        row = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb]

                        custom_status_short = \
                            df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb][
                                'custom_status_short'].values[0]
                        custom_status = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb][
                            'custom_status'].values[0]
                        decision_date = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb][
                            'decision_date'].values[0]
                        refuse_reason = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb][
                            'refuse_reason'].values[0]
                        registration_numb = df_loaded_decisions.loc[df_loaded_decisions['parcel_numb'] == parcel_numb][
                            'registration_numb'].values[0]
                        #print(parcel_numb, custom_status, custom_status_short, decision_date, registration_numb)
                        con.execute(f"Update baza set "
                                    f" registration_numb = '{registration_numb}',"
                                    f" custom_status = '{custom_status}',"
                                    f" custom_status_short = '{custom_status_short}',"
                                    f" decision_date = '{decision_date}',"
                                    f" refuse_reason = '{refuse_reason}' where parcel_numb = '{parcel_numb}'")
                        #    row_isalready_in = pd.read_sql(f"Select * from baza where parcel_numb = '{parcel_numb}'", con)
                        # df_isalready_in = df_isalready_in.append(row_isalready_in)
                    df_to_append.to_sql('baza', con=con, if_exists='append', index=False)

            except Exception as e:
                print(e)
        print(f"chunk{i} updated")


scheduler = BackgroundScheduler(daemon=True, job_defaults={'max_instances': 3})

# Create the job
scheduler.add_job(func=server_request_events, trigger='interval', seconds=500)  # trigger='cron', hour='22', minute='30'
scheduler.start()
