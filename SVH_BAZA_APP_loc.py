import base64
import hashlib
import time
import traceback
#from mysql.connector import connect, Error
import flask
import requests
from apscheduler.schedulers.background import BackgroundScheduler
from waitress import serve
from urllib.parse import urlparse
import logging
from flask import Flask, jsonify, request, render_template, redirect, url_for, send_file
from flask import abort
from flask import flash
import pandas as pd
import sqlite3 as sl
import numpy as np
import datetime
from smtplib import SMTP_SSL
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os
from flask import Flask, render_template, request, url_for
from flask_sqlalchemy import SQLAlchemy
from flask_restful import Resource, Api
from sqlalchemy.orm import sessionmaker, scoped_session
from sqlalchemy import create_engine
from sqlalchemy.ext.declarative import declarative_base
from apispec.ext.marshmallow import MarshmallowPlugin
from apispec import APISpec
from flask_apispec.extension import FlaskApiSpec
from flask_apispec import use_kwargs, marshal_with
from flask import make_response
from flask_jwt_extended import (
    JWTManager, jwt_required,
    get_jwt_identity, set_access_cookies, unset_jwt_cookies)

from SVH_BAZA_modules import manifest_services, party_to_pallet, api_views, pallet_views
from SVH_BAZA_modules.api_views import server_request_events, check_and_backup
from SVH_BAZA_modules.parcel_services import search_parcel_sql_service, add_to_zone_service, add_to_place_sql_service
from SVH_BAZA_modules.plomb_services import get_plomb_come_work_service
from schemas import UserSchema, AuthSchema
import winsound
from SVH_BAZA_modules import parcel_services
from SVH_BAZA_modules.services import (insert_user_action, map_eng_to_rus, download_folder, addition_folder,
                                       create_databases, logger, logger_change_plob, style, get_user_name)
from SVH_BAZA_modules.load_excel_service import (load_sample_manifest_service, load_decisions_service,
                                                 )

app_svh = Flask(__name__)

app_svh.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app_svh.config['UPLOAD_EXTENSIONS'] = ['.xls', '.xlsx']
app_svh.config['JWT_CSRF_CHECK_FORM'] = False
app_svh.config['JWT_TOKEN_LOCATION'] = ['cookies']
app_svh.config["SQLALCHEMY_DATABASE_URI"] = ('sqlite:///db.sqlite')

db = SQLAlchemy(app_svh)

client = app_svh.test_client()
engine = create_engine('sqlite:///db.sqlite')
session = scoped_session(sessionmaker(autocommit=False, autoflush=False, bind=engine))

Base = declarative_base()
Base.query = session.query_property()

jwt = JWTManager(app_svh)
app_svh.secret_key = 'c9e779a3258b42338334daaed51bccf7'
app_svh.config['SESSION_TYPE'] = 'filesystem'
app_svh.config['JWT_SECRET_KEY'] = 'c9e779a3258b42338334daaed51bccf7'

api = Api(app_svh)

pd.set_option("display.precision", 3)
pd.options.display.float_format = '{:.3f}'.format
con = sl.connect('BAZA.db')

app_svh.register_blueprint(party_to_pallet.bp_party_to_pallet)
app_svh.register_blueprint(api_views.bp_api)
app_svh.register_blueprint(pallet_views.bp_pallet)

now = datetime.datetime.now().strftime("%d.%m.%Y")
now_time = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")

# create all databases and tables initial
create_databases()


db_url = "mysql+mysqlconnector://{USER}:{PWD}@{HOST}/{DBNAME}"
db_url = db_url.format(
    USER="root",
    PWD="jPouKY2zy3R6",
    HOST="localhost",
    DBNAME="baza",
    auth_plugin='mysql_native_password'
)
#engine = create_engine(db_url, echo=False)

def setup_logger(name, log_file, level=logging.INFO):
    logging.basicConfig(format=u'%(levelname)-8s [%(asctime)s] %(message)s')  # filename=u'mylog.log'
    handler = logging.FileHandler(log_file)
    logger = logging.getLogger(name)
    logger.setLevel(level)
    logger.addHandler(handler)
    return logger


logger_GBS_statuses = setup_logger('logger_GBS_statuses', 'logger_GBS_statuses.log')


class UserLogin(Resource):
    def post(self):
        username = request.get_json()['username']
        password = request.get_json()['password']
        if username == 'admin' and password == 'habr':
            access_token = create_access_token(identity={
                'role': 'admin',
            }, expires_delta=False)
            result = {'token': access_token}
            return result
        return {'error': 'Invalid username and password'}


class ProtectArea(Resource):
    @jwt_required
    def get(self):
        return {'answer': 42}


api.add_resource(UserLogin, '/api/login/')
api.add_resource(ProtectArea, '/api/protect-area/')

docs = FlaskApiSpec()
docs.init_app(app_svh)
app_svh.config.update({
    'APICPEC_SPEC': APISpec(
        title='videoblog',
        version='v1',
        openapi_version='2.0',
        plugins=[MarshmallowPlugin()]
    ),
    'APISPEC_SWAGGER_URL': '/swagger/'
})

from models import *

Base.metadata.create_all(bind=engine)

app_svh.jinja_env.globals.update(get_user_name=get_user_name)

@app_svh.route('/a', methods=['GET', 'POST'])
def playAudioFile():
    return render_template('audio.html')


@app_svh.route('/todo/api/v1.0/register', methods=['POST'])
@use_kwargs(UserSchema)
@marshal_with(AuthSchema)
def register(**kwargs):
    try:
        user = User(**kwargs)
        session.add(user)
        session.commit()
        token = user.get_token()
    except Exception as e:
        logger.warning(f'register error: {e}')
        return {'message': str(e)}, 400
    return {'access_token': token}


@app_svh.route('/home')
def home():
    return render_template("home.html")


@app_svh.route('/login', methods=['GET', 'POST'])
def login_start():
    return render_template("login.html")


@app_svh.route('/get_user_id')
@jwt_required()
def get_user_id():
    user_id = get_jwt_identity()
    resp = make_response(redirect(url_for('fetchmany_party')))
    resp.set_cookie('user_id', str(user_id))
    return resp


@app_svh.route('/todo/api/v1.0/login_insert', methods=['POST'])
def login_insert():
    try:
        email = request.form['email']
        password = request.form['password']
        log = client.post('/todo/api/v1.0/login', json={'email': email, 'password': password})
        access_token = log.get_json()['access_token']
        resp = make_response(redirect(url_for('get_user_id')))
        set_access_cookies(resp, access_token)
    except Exception as e:
        logger.warning(e)
        return {'message': str(e)}, 400
    return resp


@app_svh.route('/todo/api/v1.0/login', methods=['POST'])
@use_kwargs(UserSchema(only=('email', 'password')))
@marshal_with(AuthSchema)
def login(**kwargs):
    user = User.authenticate(**kwargs)
    token = user.get_token()
    return {'access_token': token}


@app_svh.route('/logout')
@jwt_required()
def logout():
    resp = make_response(redirect(url_for('login_start')))
    # resp.set_cookie('access_token', max_age=0)
    unset_jwt_cookies(resp)
    return resp


@app_svh.teardown_appcontext
def shutdown_session(exception=None):
    session.remove()


@app_svh.errorhandler(422)
def error_handlers(err):
    headers = err.data.get('headers', None)
    messages = err.data.get('messages', ['Invalid request'])
    if headers:
        return jsonify({'message': messages}), 400, headers
    else:
        return jsonify({'message': messages}), 400


@app_svh.route('/a')
def returnAudioFile():
    path_to_audio_file = "/statiс/SndNoPlomb.wav"  # audio from project dir
    return send_file(
        path_to_audio_file,
        mimetype="audio/wav",
        as_attachment=True,
        attachment_filename="SndNoPlomb.wav")


@app_svh.route('/add/load_sample_manifest', methods=['GET', 'POST'])
def load_sample_manifest():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        filename = uploaded_file.filename
        logger.warning(filename)
        if filename != '':
            file_ext = os.path.splitext(filename)[1]
            if file_ext not in app_svh.config['UPLOAD_EXTENSIONS']:
                abort(400)
            load_sample_manifest_service(uploaded_file, filename)
    return render_template('add_manifest_sample.html')


@app_svh.route('/add/decisions', methods=['GET', 'POST'])
def load_decisions():
    logger.warning('OK')
    if request.method == 'POST':
        uploaded_file = request.files['file']
        filename = uploaded_file.filename
        logger.warning(filename)
        if filename != '':
            file_ext = os.path.splitext(filename)[1]
            if file_ext not in app_svh.config['UPLOAD_EXTENSIONS']:
                abort(400)
            load_decisions_service(uploaded_file, filename)
    return render_template('add_decisions.html')


def send_mail(filepath, subject):
    basename = os.path.basename(filepath)
    address = "logistick.dv@yandex.ru"

    # Compose attachment
    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(filepath, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="%s"' % basename)

    # Compose message
    msg = MIMEMultipart()
    msg['From'] = address
    msg['To'] = address
    msg['Subject'] = subject
    msg.attach(part)

    # Send mail
    smtp = SMTP_SSL('smtp.yandex.ru')
    smtp.connect('smtp.yandex.ru')
    smtp.login(address, 'rwlefgatbfpewlmt')
    smtp.sendmail(address, address, msg.as_string())
    smtp.quit()


df = pd.DataFrame()

@app_svh.route('/parties_analitic')
def parties_analitic():
    con = sl.connect('BAZA.db')
    con.row_factory = sl.Row
    with con:
        len_id = con.execute('SELECT max(id) from baza').fetchone()[0]
        id_for_job = len_id - 150000
        data = f"Select DISTINCT ID, party_numb from baza where ID > {id_for_job}"
        # data_df = pd.DataFrame(data)
        data_df = pd.read_sql(data, con).sort_values(by='ID', ascending=False)

        parties = data_df.drop_duplicates(subset='party_numb')['party_numb']
        parties = parties.fillna(value=np.nan)
        list_of_parties = parties.to_list()
        tuple_of_parties = tuple(list_of_parties)
        # print(tuple_of_parties)
        data2 = f'Select party_numb, parcel_plomb_numb, parcel_numb, custom_status_short, custom_status, VH_status, parcel_weight  from baza where party_numb in {tuple_of_parties}'
        parts_analit_table = pd.read_sql(data2, con)
        print(parts_analit_table)
    hostname = request.headers.get('Host')
    print(hostname)

@app_svh.route('/')
def fetchmany_party():
    con = sl.connect('BAZA.db')
    con.row_factory = sl.Row
    with con:
        len_id = con.execute('SELECT max(id) from baza').fetchone()[0]
        id_for_job = len_id - 150000
        data_start = f"Select DISTINCT ID, party_numb from baza where ID > {id_for_job}"
        # data_df = pd.DataFrame(data)
        data_df = pd.read_sql(data_start, con).sort_values(by='ID', ascending=False)
        print(data_df)
        parties = data_df.drop_duplicates(subset='party_numb')['party_numb'].dropna()

        print(parties)
        list_of_parties = parties.to_list()
        tuple_of_parties = tuple(list_of_parties)
        # print(tuple_of_parties)
        data_query = (f'Select party_numb, parcel_plomb_numb, parcel_numb, custom_status_short, '
                      f'custom_status, VH_status, parcel_weight from baza where party_numb in {tuple_of_parties}')
        data_all = pd.read_sql(data_query, con)
        print(data_all)
        analit_table = pd.DataFrame(index=None)
        for party_numb in list_of_parties:
            data = data_all.loc[(data_all['party_numb'] == party_numb)]
            a = data.drop_duplicates(subset='parcel_plomb_numb').loc[data['parcel_plomb_numb'] != '']
            df_not_shipped = a.loc[a['VH_status'] != 'ОТГРУЖЕН']
            logger.warning(a)
            quonty_plomb = len(a)
            quonty_parcels = len(data)
            not_shipt_quont = len(df_not_shipped)
            df_refuses = data.loc[(data['custom_status_short'] == 'ИЗЪЯТИЕ')]
            quonty_parcels_refuse = len(df_refuses)
            quonty_plomb_refuse = len(
                df_refuses.loc[data['parcel_plomb_numb'] != ''].drop_duplicates(subset='parcel_plomb_numb'))
            custom_control = data.loc[data['custom_status'].str.contains('родление')]
            custom_control_quont = len(custom_control)
            dont_declarate = data.loc[data['custom_status'].str.contains('Unknown')]
            dont_declarate = dont_declarate[['parcel_numb', 'parcel_plomb_numb']]
            dont_declarate_quont = len(dont_declarate)
            weight = round(data['parcel_weight'].sum(), 3)
            df_to_append = pd.DataFrame(
                data={"Партия": [party_numb], "Кол-во пломб": [quonty_plomb], "Кол-во посылок": [quonty_parcels],
                      "Отказных пломб": [quonty_plomb_refuse], "Отказных посылок": [quonty_parcels_refuse],
                      "вес партии": [weight], "Не отгруженные места": [not_shipt_quont]}, index=None)
            analit_table = analit_table.append(df_to_append)
        analit_table = analit_table.reset_index(drop=True)
        analit_table = analit_table.transpose()
        print(analit_table)
    con.close()
    return render_template('index.html', parties=parties,
                           analit_table=analit_table)


@app_svh.route('/all_parties')
def all_parties():
    con = sl.connect('BAZA.db')
    con.row_factory = sl.Row
    with con:
        data = pd.read_sql("Select DISTINCT ID, party_numb from baza", con)
        data = data.sort_values(by='ID', ascending=False)
        parties = data.drop_duplicates(subset='party_numb')['party_numb']
        parties = parties.fillna(value=np.nan)
        logger.warning(parties)
    con.close()
    return render_template('index2.html', parties=parties)


@app_svh.route('/info/not_shiped', methods=['GET', 'POST'])
def all_not_shipped():
    global df_not_shipped
    con = sl.connect('BAZA.db')
    with con:
        df_not_shipped = pd.read_sql("Select * from baza where VH_status != 'ОТГРУЖЕН'",
                                     con)
        df_not_shipped['decision_date'] = df_not_shipped['decision_date'].str.slice(0, 17)
        df_not_shipped = df_not_shipped.rename(columns=map_eng_to_rus)
        df_not_shipped['Товары'] = df_not_shipped['Товары'].str.slice(0, 50)

    return render_template('info_not_shipped.html', tables=[style + df_not_shipped.to_html(index=False,
                                                                                           float_format='{:2,.2f}'.format)],
                           titles=[''])


@app_svh.route('/info/not_shiped_to_xl', methods=['GET', 'POST'])
def not_shiped_to_xl():
    global df_not_shipped
    now_time = datetime.datetime.now().strftime("%d.%m.%Y %H-%M")
    writer = pd.ExcelWriter(f'{download_folder}Не отгруженное на {now_time}.xlsx', engine='xlsxwriter')
    df_not_shipped.to_excel(writer, sheet_name='Sheet1', index=False)
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
    flash(f'Не отгруженные список выгружен в excel!', category='success')
    object_name = ''
    comment = f'Не отгруженные список выгружен в excel!'
    insert_user_action(object_name, comment)
    return render_template('info_not_shipped.html', tables=[style + df_not_shipped.to_html(index=False,
                                                                                           float_format='{:2,.2f}'.format)],
                           titles=[''])


@app_svh.route('/info', methods=['GET', 'POST'])
def object_info():
    numb = request.form['numb']
    con = sl.connect('BAZA.db')
    with con:
        df = pd.read_sql(f"Select * from baza where parcel_numb = '{numb}'", con)
        df['decision_date'] = df['decision_date'].str.slice(0, 17)
        df = df.rename(columns=map_eng_to_rus)
        if df['Статус ТО (кратк)'].str.contains("ИЗЪЯТИЕ").any():
            try:
                if 'Требуется' in df['Статус ТО'].values[0] or 'уплат' in \
                        df['Причина отказа'].values[0] or 'Не уплачены' in \
                        df['Причина отказа'].values[0]:
                    pay_trigger = 'ПЛАТНАЯ'
                else:
                    pay_trigger = ''
            except:
                pay_trigger = ''
            flash(f'ИЗЪЯТИЕ на склад {pay_trigger}', category='error')
        elif df['Статус ТО (кратк)'].str.contains("ВЫПУСК").any():
            flash(f'ВЫПУСК!', category='success')
        else:
            pass

        df = df.transpose()
        index = True
        df = df.rename(columns={0: ''})
        object_name = 'Посылка'
        try:
            con_vect = sl.connect('VECTORS.db')
            with con_vect:
                vector = con_vect.execute(f"SELECT vector FROM VECTORS where parcel_numb = '{numb}'").fetchone()[0]
                flash(f'{vector}!', category='success')
        except Exception as e:
            vector = None
            print(e)
            pass
        if df.empty:
            df = pd.read_sql(f"Select * from baza where party_numb = '{numb}'", con)
            df['decision_date'] = df['decision_date'].str.slice(0, 17)
            df = df.drop_duplicates(subset='parcel_plomb_numb')
            df = df.rename(columns=map_eng_to_rus)
            df = df[['Партия', 'Пломба', 'Паллет', 'Зона']]
            object_name = 'Партия'
            index = False
            if df.empty:
                df = pd.read_sql(f"Select * from baza where registration_numb = '{numb}'", con)
                df['decision_date'] = df['decision_date'].str.slice(0, 17)
                df = df.rename(columns=map_eng_to_rus)
                df['Товары'] = df['Товары'].str.slice(0, 50)
                object_name = 'Реестр ПТДЭГ'
                index = False
                if df.empty:
                    df = pd.read_sql(f"Select * from baza where pallet = '{numb}'", con)
                    df['decision_date'] = df['decision_date'].str.slice(0, 17)
                    df = df.drop_duplicates(subset='parcel_plomb_numb')
                    df = df.rename(columns=map_eng_to_rus)
                    df['№'] = np.arange(len(df))[::+1] + 1
                    df = df[['№', 'Пломба', 'Зона']]
                    object_name = 'Паллет'
                    index = False
                    if df.empty:
                        df = pd.read_sql(f"Select * from baza where parcel_plomb_numb = '{numb}' COLLATE NOCASE", con)
                        df['decision_date'] = df['decision_date'].str.slice(0, 17)
                        df = df.rename(columns=map_eng_to_rus)
                        df['Товары'] = df['Товары'].str.slice(0, 50)
                        object_name = 'Пломба'
                        index = False
                        try:
                            con_vect = sl.connect('VECTORS.db')
                            with con_vect:
                                vector = con_vect.execute(
                                    f"SELECT vector FROM VECTORS where parcel_plomb_numb = '{numb}'").fetchone()[0]
                                try:
                                    flash(f'{vector}!', category='success')
                                    print(vector)
                                except:
                                    try:
                                        with con:
                                            data = con.execute(
                                                f"SELECT parcel_numb from baza where where parcel_plomb_numb = '{numb}'").fetchone()[
                                                0]
                                            try:
                                                first_parcel = data
                                                vector = con_vect.execute(
                                                    f"SELECT vector FROM VECTORS where parcel_numb = '{first_parcel}'").fetchone()[
                                                    0]
                                                try:
                                                    flash(f'{vector}!', category='success')
                                                except Exception as e:
                                                    print(e)
                                            except Exception as e:

                                                print(e)
                                    except Exception as e:
                                        print(e)
                        except Exception as e:
                            print(e)
                            pass

    # create an output stream
    return render_template('info.html', tables=[style + df.to_html(classes='mystyle', index=index,
                                                                   float_format='{:2,.2f}'.format)],
                           titles=['Информация'],
                           numb=numb, object_name=object_name, vector=vector)


@app_svh.route('/party_info/<string:row>', methods=['GET', 'POST'])
def party_info(row):
    con = sl.connect('BAZA.db')
    data = pd.read_sql(f"Select * from baza where party_numb = '{row}'", con)
    a = data.drop_duplicates(subset='parcel_plomb_numb').loc[data['parcel_plomb_numb'] != '']
    logger.warning(a)
    quonty_plomb = len(a)
    quonty_parcels = len(data)
    df_refuses = data.loc[data['custom_status_short'] == 'ИЗЪЯТИЕ']
    quonty_parcels_refuse = len(df_refuses)
    quonty_plomb_refuse = len(
        df_refuses.loc[data['parcel_plomb_numb'] != ''].drop_duplicates(subset='parcel_plomb_numb'))
    print(quonty_plomb_refuse)
    custom_control = data.loc[data['custom_status'].str.contains('родление')]
    custom_control = custom_control[['registration_numb', 'parcel_numb', 'goods']]
    custom_control = custom_control.rename(
        columns={'registration_numb': 'Рег. номер', 'parcel_numb': 'Трек', 'goods': 'Товары'})
    custom_control_quont = len(custom_control)
    dont_declarate = data.loc[data['custom_status'].str.contains('Unknown')]
    dont_declarate = dont_declarate[['registration_numb', 'parcel_numb', 'parcel_plomb_numb', 'goods']]
    dont_declarate = dont_declarate.rename(
        columns={'registration_numb': 'Рег. номер', 'parcel_numb': 'Трек', 'parcel_plomb_numb': 'Пломба',
                 'goods': 'Товары'})
    dont_declarate_quont = len(dont_declarate)
    return render_template('party_info.html', row=row,
                           quonty_plomb=quonty_plomb,
                           quonty_plomb_refuse=quonty_plomb_refuse,
                           quonty_parcels=quonty_parcels,
                           quonty_parcels_refuse=quonty_parcels_refuse,
                           custom_control=custom_control,
                           custom_control_quont=custom_control_quont,
                           dont_declarate_quont=dont_declarate_quont,
                           df_refuses=df_refuses,
                           tables=[style + custom_control.to_html(index=False,
                                                                  float_format='{:2,.2f}'.format),
                                   style + dont_declarate.to_html(index=False,
                                                                  float_format='{:2,.2f}'.format)
                                   ],
                           titles=['Информация о посылке', f'Продление срока выпуска {custom_control_quont}шт:',
                                   f'Неподанные: {dont_declarate_quont}шт']
                           )


@app_svh.route('/party_info_create_pallet/<string:row>', methods=['GET', 'POST'])
def party_info_create_pallet(row):
    return render_template('party_info_create_pallet.html', row=row, )


@app_svh.route('/party_info_refuses/<string:row>', methods=['GET', 'POST'])
def party_info_refuses(row):
    con = sl.connect('BAZA.db')
    data = pd.read_sql(f"Select * from baza where party_numb = '{row}'", con)
    a = data.drop_duplicates(subset='parcel_plomb_numb').loc[data['parcel_plomb_numb'] != '']
    logger.warning(a)
    quonty_plomb = len(a)
    quonty_parcels = len(data)
    df_refuses = data.loc[data['custom_status_short'] == 'ИЗЪЯТИЕ']
    quonty_parcels_refuse = len(df_refuses)
    quonty_plomb_refuse = len(
        df_refuses.loc[data['parcel_plomb_numb'] != ''].drop_duplicates(subset='parcel_plomb_numb'))
    df_refuses = df_refuses.rename(columns=map_eng_to_rus)
    df_refuses['№'] = np.arange(len(df_refuses))[::+1] + 1
    df_refuses = df_refuses[
        ['№', 'Партия', 'Трек-номер', 'Пломба', 'Статус ТО', 'Статус ТО (кратк)', 'Дата решения', 'Причина отказа',
         'вес',
         'Товары', 'Паллет']]
    return render_template('party_info_refuses.html', row=row,
                           quonty_plomb=quonty_plomb,
                           quonty_plomb_refuse=quonty_plomb_refuse,
                           quonty_parcels=quonty_parcels,
                           quonty_parcels_refuse=quonty_parcels_refuse,
                           tables=[style + df_refuses.to_html(index=False,
                                                              float_format='{:2,.2f}'.format)],
                           titles=['', f'Список изъятий по партии {row}']
                           )


@app_svh.route('/party_info_refuses_excel/<string:row>', methods=['GET', 'POST'])
def party_info_refuses_excel(row):
    con = sl.connect('BAZA.db')
    data = pd.read_sql(f"Select * from baza where party_numb = '{row}'", con)
    a = data.drop_duplicates(subset='parcel_plomb_numb').loc[data['parcel_plomb_numb'] != '']
    quonty_plomb = len(a)
    quonty_parcels = len(data)
    df_refuses = data.loc[data['custom_status_short'] == 'ИЗЪЯТИЕ']
    quonty_parcels_refuse = len(df_refuses)
    quonty_plomb_refuse = len(
        df_refuses.loc[data['parcel_plomb_numb'] != ''].drop_duplicates(subset='parcel_plomb_numb'))
    df_refuses = df_refuses.rename(columns=map_eng_to_rus)
    writer = pd.ExcelWriter(f'{download_folder}ИЗЪЯТИЯ {row}.xlsx', engine='xlsxwriter')
    df_refuses.to_excel(writer, sheet_name='Sheet1', index=False)
    for column in df_refuses:
        column_width = max(df_refuses[column].astype(str).map(len).max(), len(column))
        col_idx = df_refuses.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
        writer.sheets['Sheet1'].set_column(0, 3, 10)
        writer.sheets['Sheet1'].set_column(1, 3, 20)
        writer.sheets['Sheet1'].set_column(2, 3, 20)
        writer.sheets['Sheet1'].set_column(3, 3, 20)
        writer.sheets['Sheet1'].set_column(4, 3, 30)
        writer.sheets['Sheet1'].set_column(5, 3, 20)
    writer.save()
    flash(f'Изъятия по партии {row} выгружены в excel!', category='success')
    object_name = row
    comment = f'Изъятия по партии {row} выгружены в excel!'
    user_id = request.cookies.get('user_id')
    con_user = sl.connect("db.sqlite")
    query = f"Select name from users where id = {user_id}"
    user_name = con_user.execute(query).fetchone()
    print(user_name)
    insert_user_action(object_name, comment)
    return render_template('party_info_refuses.html', row=row,
                           quonty_plomb=quonty_plomb,
                           quonty_plomb_refuse=quonty_plomb_refuse,
                           quonty_parcels=quonty_parcels,
                           quonty_parcels_refuse=quonty_parcels_refuse, user_id=user_id)


@app_svh.route('/party_info_vectors/<string:row>', methods=['GET', 'POST'])
def party_info_vectors(row):
    con = sl.connect('BAZA.db')
    con_vect = sl.connect('VECTORS.db')
    print(row)
    with con_vect:
        df_vectors = pd.read_sql(f"Select * from vectors where party_numb = '{row}'", con_vect)
    print(df_vectors)
    a = df_vectors.drop_duplicates(subset='parcel_plomb_numb').loc[df_vectors['parcel_plomb_numb'] != '']
    logger.warning(a)
    quonty_plomb = len(a)
    quonty_parcels = len(df_vectors)
    df_vectors['№'] = np.arange(len(df_vectors))[::+1] + 1
    df_vectors['quont'] = 1
    print(df_vectors)
    group_df_vectors = df_vectors.drop_duplicates(subset='parcel_plomb_numb').loc[df_vectors['parcel_plomb_numb'] != '']
    with con:
        query = f"Select DISTINCT parcel_plomb_numb from baza where party_numb = '{row}' and VH_status is ?"
        df_not_shiped = pd.read_sql(sql=query, con=con, params=(None,))
        print(df_not_shiped)
        df_not_shiped['quont2'] = 1
        df_not_shiped['quont2'] = df_not_shiped['quont2'].astype(int)
        pd.options.display.float_format = '{:,.0f}'.format
    group_df_vectors_all = pd.merge(group_df_vectors, df_not_shiped, how='left', left_on='parcel_plomb_numb',
                                    right_on='parcel_plomb_numb')

    group_df_vectors = group_df_vectors_all.groupby('vector')['quont', 'quont2'].sum().reset_index()  # .to_frame()
    print(group_df_vectors)
    group_df_vectors = group_df_vectors.rename(
        columns={'vector': 'Направление', 'quont': 'Кол-во пломб', 'quont2': 'Не отгруженны'})
    group_df_vectors_all = group_df_vectors_all.loc[group_df_vectors_all['quont2'] == 1]
    group_df_vectors_all['№п/п'] = np.arange(len(group_df_vectors_all))[::+1] + 1
    group_df_vectors_all = group_df_vectors_all.rename(columns={'parcel_plomb_numb': 'Пломба',
                                                                'vector': 'Направление'})
    group_df_vectors_all = group_df_vectors_all.reindex(columns=['№п/п',
                                                                 'Пломба',
                                                                 'Направление'])
    return render_template('party_info_vectors.html', row=row,
                           quonty_plomb=quonty_plomb,
                           quonty_parcels=quonty_parcels,
                           group_df_vectors_all=group_df_vectors_all,
                           tables=[style + group_df_vectors.to_html(index=False,
                                                                    float_format='{:,.0f}'.format),
                                   style + group_df_vectors_all.to_html(index=False,
                                                                        float_format='{:,.0f}'.format)
                                   ],
                           titles=['', f'Направления по партии {row}', 'Неотгруженные пломбы:']
                           )


@app_svh.route('/check_refuses/<string:row>', methods=['GET', 'POST'])
def check_refuses(row):
    party_numb = row
    con = sl.connect('BAZA.db')
    with con:
        try:
            con = sl.connect('BAZA.db')
            with con:
                baza = con.execute("select count(*) from sqlite_master where type='table' and name='party_refuses'")
                for row in baza:
                    # если таких таблиц нет
                    if row[0] == 0:
                        # создаём таблицу
                        with con:
                            con.execute("""
                                                    CREATE TABLE party_refuses (
                                                    ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                    party_numb VARCHAR(20),
                                                    parcel_numb VARCHAR(20),
                                                    custom_status_short VARCHAR(8),
                                                    parcel_find_status VARCHAR(8)
                                                    );
                                                """)

                df = pd.read_sql(
                    f"SELECT party_numb, parcel_numb, custom_status_short FROM baza where party_numb = '{party_numb}' and custom_status_short = 'ИЗЪЯТИЕ' COLLATE NOCASE",
                    con)
                df['parcel_find_status'] = " "
                df['№'] = np.arange(len(df))[::+1] + 1
                df = df[['№', 'party_numb', 'parcel_numb', 'parcel_find_status']]
                index = False
                print(df)
                df.to_sql('party_refuses', con=con, if_exists='replace', index=False)
        except Exception as e:
            return {'message': str(e)}, 400
    return render_template('party_refuses_work.html', tables=[style + df.to_html(classes='mystyle', index=index,
                                                                                 float_format='{:2,.2f}'.format)],
                           titles=['Информация'],
                           party_numb=party_numb)


@app_svh.route('/check_refuses_work/<string:party_numb>', methods=['GET', 'POST'])
def check_refuses_work(party_numb):
    parcel_numb = request.form['parcel_numb']
    con = sl.connect('BAZA.db')
    with con:
        df_check_parcel = pd.read_sql(
            f"SELECT * FROM party_refuses where parcel_numb = '{parcel_numb}' COLLATE NOCASE", con)
        if df_check_parcel.empty:
            flash(f'Посылка {parcel_numb} не входит в список изъятий или не найдена!', category='error')
            winsound.PlaySound('Snd\Snd_CancelIssue.wav', winsound.SND_FILENAME)
        else:
            con.execute(
                f"Update party_refuses set parcel_find_status = 'НАЙДЕНА' where parcel_numb = '{parcel_numb}'")

            df_parc_events = pd.read_sql(f"SELECT * FROM baza where parcel_numb = '{parcel_numb}'", con)
            try:
                if 'Требуется' in df_parc_events['custom_status'].values[0] or 'уплат' in \
                        df_parc_events['refuse_reason'].values[0] or 'Не уплачены' in df_parc_events['refuse_reason'].values[0]:
                    pay_trigger = 'ПЛАТНАЯ'
                    flash(f'{pay_trigger}', category='error')
                else:
                    pay_trigger = ''
            except:
                pay_trigger = ''
        df = pd.read_sql(f"SELECT * FROM party_refuses where party_numb = '{party_numb}' COLLATE NOCASE", con)
        df['№'] = np.arange(len(df))[::+1] + 1
        df = df[['№', 'party_numb', 'parcel_numb', 'parcel_find_status']]
        if df.loc[df['parcel_find_status'] == ' '].empty:
            flash(f'Все изъятия найдены', category='success')
            winsound.PlaySound('Snd\se_mesta_naid.wav', winsound.SND_FILENAME)
        index = False
        quont_refuses_all = len(df)
        quont_refuses_done = len(df.loc[df['parcel_find_status'] == 'НАЙДЕНА'])
        quon_not_done = quont_refuses_all - quont_refuses_done
        object_name = parcel_numb
        comment = f'Проверка изъятий по партии {party_numb}: Посылка просмотренна'
        insert_user_action(object_name, comment)

    return render_template('party_refuses_work2.html', tables=[style + df.to_html(classes='mystyle', index=index,
                                                                                  float_format='{:2,.2f}'.format)],
                           titles=['Информация'],
                           party_numb=party_numb, quont_refuses_all=quont_refuses_all,
                           quont_refuses_done=quont_refuses_done, quon_not_done=quon_not_done)


@app_svh.route('/add/load_tracks', methods=['GET', 'POST'])
def load_tracks():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        filename = uploaded_file.filename
        logger.warning(filename)
        if filename != '':
            file_ext = os.path.splitext(filename)[1]
            if file_ext not in app_svh.config['UPLOAD_EXTENSIONS']:
                abort(400)
            uploaded_file.save(uploaded_file.filename)
            df_track = pd.read_excel(filename, sheet_name=0, engine='openpyxl', usecols='A, AO')
            df_track = df_track[['Номер отправления ИМ', 'Номер накладной СДЭК']]
            df_track = df_track.rename(columns={'Номер отправления ИМ': 'parcel_numb',
                                                'Номер накладной СДЭК': 'track_numb'})
            if df_track['track_numb'].str.contains('#н/д').any() or df_track['track_numb'].isnull().any():
                flash(
                    f'Ошибка загрузки треков: В колонке Треков есть пустые значения или #н/д, поправьте и загрузите заново',
                    category='error')
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
                flash(f'Треки загружены')
                winsound.PlaySound('Snd\sample_load.wav', winsound.SND_FILENAME)

    return render_template('add_manifest_sample.html')


@app_svh.route('/add/vectors', methods=['GET', 'POST'])
def load_vectors():
    if request.method == 'POST':
        uploaded_file = request.files['file']
        filename = uploaded_file.filename
        logger.warning(filename)
        if filename != '':
            file_ext = os.path.splitext(filename)[1]
            if file_ext not in app_svh.config['UPLOAD_EXTENSIONS']:
                abort(400)
            uploaded_file.save(uploaded_file.filename)
            df_vector = pd.read_excel(filename, sheet_name=0, engine='openpyxl')
            df_vector = df_vector[['Партия', 'Номер пломбы', 'Номер отправления ИМ', 'Направление']]
            df_vector = df_vector.rename(columns={'Партия': 'party_numb', 'Номер пломбы': 'parcel_plomb_numb',
                                                  'Номер отправления ИМ': 'parcel_numb',
                                                  'Направление': 'vector'})
            if df_vector['vector'].str.contains('#н/д').any() or df_vector['parcel_numb'].isnull().any():
                flash(
                    f'Ошибка загрузки Направлений: В колонке направлений есть пустые значения или #н/д, поправьте и загрузите заново',
                    category='error')
            else:
                con_vector = sl.connect('VECTORS.db')
                with con_vector:
                    data = con_vector.execute(
                        "select count(*) from sqlite_master where type='table' and name='vectors'")
                    for row in data:
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
                    df_vector.to_sql('vectors', con=con_vector, if_exists='append', index=False)
                    con_vector.commit()
                flash(f'Направления загружены')
                winsound.PlaySound('Snd\sample_load.wav', winsound.SND_FILENAME)

    return render_template('add_manifest_sample.html')


@app_svh.route('/plomb', methods=['GET'])
@jwt_required()
def plomb():
    return render_template('plomb.html')


@app_svh.route('/pallet', methods=['GET'])
@jwt_required()
def pallet():
    return render_template('pallet.html')


@app_svh.route('/search/plomb', methods=['GET'])
@jwt_required()
def plomb_searh():
    try:
        parcel_plomb_numb = request.args.get('parcel_plomb_numb')
        return render_template('plomb_search.html', search=parcel_plomb_numb)
    except Exception as e:
        return {'message': str(e)}, 400


@app_svh.route('/search/parcel', methods=['GET'])
def parcel_searh():
    try:
        parcel_numb = request.args.get('parcel_numb')
        return render_template('parcel_search.html', search=parcel_numb)
    except Exception as e:
        return {'message': str(e)}, 400


@app_svh.route('/search/parcel_goods_info', methods=['GET'])
def parcel_goods_info():
    try:
        parcel_numb = request.args.get('parcel_numb')
        return render_template('parcel_goods.html', search=parcel_numb)
    except Exception as e:
        return {'message': str(e)}, 400


@app_svh.route('/search/plomb_info', methods=['POST', 'GET'])
def get_plomb_info():
    parcel_plomb_numb = request.form['parcel_plomb_numb']
    title_status = ''
    vector = None
    try:
        con = sl.connect('BAZA.db')
        with con:
            data = con.execute('SELECT * FROM plombs WHERE parcel_plomb_numb=?', (parcel_plomb_numb,)).fetchone()
            try:
                if data is None:
                    pass
                else:
                    con.execute(
                        f"Update plombs set parcel_plomb_status = 'Принят' where parcel_plomb_numb = '{parcel_plomb_numb}'")
            except:
                pass
            df_search_plomb = pd.read_sql(
                f"SELECT * FROM baza where parcel_plomb_numb = '{parcel_plomb_numb}' COLLATE NOCASE", con)
            df_search_plomb['№'] = np.arange(len(df_search_plomb))[::+1] + 1
            first_parcel = df_search_plomb['parcel_numb'].values[0]
            print(first_parcel)
            con_vect = sl.connect('VECTORS.db')
            with con_vect:
                try:
                    vector = con_vect.execute(
                        f"SELECT vector FROM VECTORS where parcel_plomb_numb = '{parcel_plomb_numb}'").fetchone()[0]
                    flash(f'{vector}!', category='success')
                    print(vector)
                except:
                    try:
                        vector = con_vect.execute(
                            f"SELECT vector FROM VECTORS where parcel_numb = '{first_parcel}'").fetchone()[
                            0]
                        flash(f'{vector}!', category='success')
                        print(vector)
                    except Exception as e:
                        print(e)
            try:
                df_search_plomb['parcel_weight'] = df_search_plomb['parcel_weight'].round(3)
            except:
                pass
            df_search_plomb = df_search_plomb[
                ['№', 'parcel_numb', 'custom_status_short', 'VH_status', 'parcel_weight', 'goods']]
            df_search_plomb = df_search_plomb.rename(columns={'parcel_numb': 'Трек-номер',
                                                              'custom_status_short': 'Статус',
                                                              'parcel_weight': 'Вес',
                                                              'VH_status': 'ВХ', 'goods': 'Товары'})
            df_search_plomb['Товары'] = df_search_plomb['Товары'].str.slice(0, 50)
            df_parc_quont = len(df_search_plomb)
            df_parc_refuse_quont = len(df_search_plomb.loc[df_search_plomb['Статус'] == 'ИЗЪЯТИЕ'])
            if df_search_plomb.empty:
                flash(f'Пломба не найдена!', category='error')
                winsound.PlaySound('Snd\Snd_NoPlomb.wav', winsound.SND_FILENAME)
            elif df_search_plomb['Статус'].str.contains("ИЗЪЯТИЕ").any():
                flash(f'Открываем место!', category='error')
                winsound.PlaySound('Snd\Snd_Open_Bag.wav', winsound.SND_FILENAME)
                title_status = 'ИЗЪЯТИЕ'
            else:
                title_status = 'ВЫПУСК'
            df_search_plomb.fillna("", inplace=True)

            def highlight_last_row(df_search_plomb):
                return ['background-color: #FF0000' if 'ИЗЪЯТИЕ' in str(i) else '' for i in df_search_plomb]

            def highlight_last_row_2(df_search_plomb):
                return ['background-color: #7CFC00' if 'На ВХ' in str(i) else '' for i in df_search_plomb]

            if 'На ВХ' in df_search_plomb['ВХ'].values:
                df_search_plomb = df_search_plomb.style.apply(highlight_last_row_2).hide_index()
                logger.warning('VH')
            else:
                df_search_plomb = df_search_plomb.style.apply(highlight_last_row).hide_index()
            df_search_plomb = df_search_plomb.format(precision=2)
        object_name = parcel_plomb_numb
        logger.warning(object_name)

        comment = 'Отбор пломб: Просмотрена пломба'
        insert_user_action(object_name, comment)
    except Exception as e:
        winsound.PlaySound('Snd\Snd_NoPlomb.wav', winsound.SND_FILENAME)
        flash(f'Пломба не найдена! {e}', category='error')
        return render_template('plomb_info.html')
    user_id, user_name = get_user_name()
    return render_template('plomb_info.html', tables=[style + df_search_plomb.to_html(index=False,
                                                                                      float_format='{:2,.2f}'.format)],
                           titles=['Информация о посылке'],
                           parcel_plomb_numb=parcel_plomb_numb, df_parc_quont=df_parc_quont,
                           df_parc_refuse_quont=df_parc_refuse_quont, title_status=title_status,
                           vector=vector, user_id=user_id, user_name=user_name)


@app_svh.route('/party/plomb_come/<party_numb>', methods=['POST', 'GET'])
def get_plomb_come(party_numb):
    try:
        con = sl.connect('BAZA.db')
        with con:
            baza = con.execute("select count(*) from sqlite_master where type='table' and name='plombs'")
            for row in baza:
                # если таких таблиц нет
                if row[0] == 0:
                    # создаём таблицу
                    with con:
                        con.execute("""
                                                CREATE TABLE plombs (
                                                ID INTEGER PRIMARY KEY AUTOINCREMENT,
                                                party_numb VARCHAR(20),
                                                parcel_plomb_numb VARCHAR(20),
                                                parcel_plomb_status VARCHAR(8)
                                                );
                                            """)
            row_isalready_in_plombs = pd.read_sql(f"Select * from plombs where party_numb = '{party_numb}'", con)
            if row_isalready_in_plombs.empty:
                df = pd.read_sql(f"SELECT * FROM baza where party_numb = '{party_numb}' COLLATE NOCASE", con)
                df = df.drop_duplicates(subset='parcel_plomb_numb')
                df['parcel_plomb_status'] = 'Ожидаем'
                df = df[['parcel_plomb_numb', 'party_numb', 'parcel_plomb_status']]
                index = False
                print(df)
                df.to_sql('plombs', con=con, if_exists='append', index=False)
            else:
                df = pd.read_sql(f"SELECT * FROM plombs where party_numb = '{party_numb}' COLLATE NOCASE", con)
                df = df.drop_duplicates(subset='parcel_plomb_numb').loc[df['parcel_plomb_numb'] != '']
                df['№'] = np.arange(len(df))[::+1] + 1
                df = df[['№', 'parcel_plomb_numb', 'party_numb', 'parcel_plomb_status']]
                index = False
                print(df, 'j')
            object_name = party_numb
            comment = f'Приемка по местам: Открыта приемка по партии {party_numb})'
            insert_user_action(object_name, comment)
    except Exception as e:
        return {'message': str(f'{e}')}, 400
    return render_template('party_plombs.html', tables=[style + df.to_html(classes='mystyle', index=index,
                                                                           float_format='{:2,.2f}'.format)],
                           titles=['Информация'],
                           party_numb=party_numb)


@app_svh.route('/party/plomb_come_work/<party_numb>', methods=['POST', 'GET'])
def get_plomb_come_work(party_numb):
    (index, quont_all_plombs, quont_plomb_done, quon_not_done,
     vector, df, parcel_plomb_numb) = get_plomb_come_work_service(party_numb)
    return render_template('party_plombs_work.html', tables=[style + df.to_html(classes='mystyle', index=index,
                                                                                float_format='{:2,.2f}'.format)],
                           titles=['Информация'],
                           party_numb=party_numb, quont_all_plombs=quont_all_plombs,
                           quont_plomb_done=quont_plomb_done, quon_not_done=quon_not_done,
                           parcel_plomb_numb=parcel_plomb_numb,
                           vector=vector)


done_parcels = pd.DataFrame()


@app_svh.route('/search/parcel_info', methods=['POST', 'GET'])
def get_parcel_info():
    global done_parcels
    try:
        (df_parcel_plomb_refuse_info, done_parcels_styl, parcel_numb, audiofile,
         df_parc_quont, df_parc_refuse_quont,
         done_parcels, parcel_plomb_numb) = parcel_services.get_parcel_info_service(done_parcels)
        return render_template('parcel_info.html', tables=[
            df_parcel_plomb_refuse_info.to_html(classes='mystyle',
                                                index=False,
                                                float_format='{:2,.2f}'.format),
            done_parcels_styl.to_html(classes='mystyle', index=False,
                                      float_format='{:2,.2f}'.format)],
                               titles=['na', 'Нужно найти в мешке:', '\n\nОтработанные:'],
                               parcel_numb=parcel_numb, df_parc_quont=df_parc_quont,
                               df_parc_refuse_quont=df_parc_refuse_quont, parcel_plomb_numb=parcel_plomb_numb,
                               audiofile=audiofile)

    except Exception as e:
        flash(f'Посылка не найдена! {e}', category='error')
        # winsound.PlaySound('Snd\Snd_Parcel_Not_Found.wav', winsound.SND_FILENAME)
        audiofile = 'Snd_CancelIssue.wav'
        return render_template('parcel_info.html', audiofile=audiofile)


@app_svh.route('/search/parcel_info_sql_first', methods=['POST', 'GET'])
def get_parcel_info_sql_first():
    try:
        user_name, user_id = get_user_name()
        con = sl.connect("BAZA.db")
        parcel_plomb_numb = request.form['parcel_plomb_numb']
        with con:
            df_plomb = pd.read_sql(
                f"SELECT parcel_numb, parcel_plomb_numb, custom_status_short, goods, parcel_weight FROM baza where parcel_plomb_numb = '{parcel_plomb_numb}'",
                con)
            if not df_plomb.empty:
                df_plomb['user_id'] = user_id
                print(df_plomb)
                df = pd.read_sql(f"Select * from parcels_refuses where parcel_plomb_numb = '{parcel_plomb_numb}'", con)
                print(df)
                if df.empty:
                    df_plomb.to_sql("parcels_refuses", con, if_exists='append', index=False)
                else:
                    for parcel_numb in df['parcel_numb']:
                        custom_status_short = df_plomb.loc[df_plomb['parcel_numb'] == parcel_numb]['custom_status_short'].values[0]
                        con.execute(f"Update parcels_refuses set user_id = {user_id}, custom_status_short = '{custom_status_short}' where parcel_numb = '{parcel_numb}'")
                df = pd.read_sql(
                        f"SELECT user_id, parcel_numb, parcel_plomb_numb, custom_status_short, parcel_find_status, goods, parcel_weight FROM parcels_refuses where parcel_plomb_numb = '{parcel_plomb_numb}'",
                        con)
                print(df)
                df_refuses = df.loc[df['custom_status_short'] == "ИЗЪЯТИЕ"].fillna('')
                qt_refuse = len(df_refuses)
                qt_all = len(df_plomb)
                qt_found = len(df_refuses.loc[df_refuses['parcel_find_status'] == 'НАЙДЕНА'])

                if qt_refuse == qt_found:
                    flash(f'Все отказы найдены!', category='success')
                df_refuses['№'] = np.arange(len(df_refuses))[::+1] + 1
                df_refuses = df_refuses[
                    ['№', 'parcel_numb', 'parcel_find_status', 'custom_status_short', 'goods', 'user_id']]
                print(df_refuses)
                object_name = parcel_plomb_numb
                comment = f'Отбор посылок: Начат отбор по пломбе'
                insert_user_action(object_name, comment)
                if df_refuses.empty:
                    flash(f'Все место выпущено!', category='success')
                    audiofile = 'Snd_All_Issue.wav'
                    return render_template('parcel_search.html', audiofile=audiofile)
                else:
                    audiofile = 'Snd_CancelIssue.wav'
                    return render_template('parcel_info_sql.html',
                                           parcel_plomb_numb=parcel_plomb_numb, qt_refuse=qt_refuse, qt_all=qt_all, qt_found=qt_found,
                                           tables=[style + df_refuses.to_html(classes='mystyle', index=False,
                                                                                            float_format='{:2,.2f}'.format)],
                                           audiofile=audiofile,
                                            titles=['Информация'])
            else:
                flash(f'Место {parcel_plomb_numb} не найдено!', category='error')
        return render_template('parcel_search.html',
                               parcel_plomb_numb=parcel_plomb_numb)
    except Exception as e:
        flash(f'Место не найдено! ошибка: {e}', category='error')
    return render_template('parcel_search.html')


@app_svh.route('/search/parcel_info_sql', methods=['POST', 'GET'])
def get_parcel_info_sql():
    try:
        parcel_numb, parcel_plomb_numb, audiofile, df_refuses, qt_refuse, qt_all, qt_found = search_parcel_sql_service()
        print(parcel_plomb_numb)
        return render_template('parcel_info_sql.html', parcel_numb=parcel_numb,
                           parcel_plomb_numb=parcel_plomb_numb, qt_refuse=qt_refuse, qt_all=qt_all, qt_found=qt_found,
                           audiofile=audiofile, tables=[style + df_refuses.to_html(classes='mystyle', index=False,
                                                                                float_format='{:2,.2f}'.format)],
                           titles=['Информация'])
    except Exception as e:
        print(e)
        flash(f'Посылка не найдена!', category='error')
        # winsound.PlaySound('Snd\Snd_Parcel_Not_Found.wav', winsound.SND_FILENAME)
        audiofile = 'Snd_CancelIssue.wav'
        return render_template('parcel_info_sql.html', audiofile=audiofile)

@app_svh.route('/search/clean', methods=['POST'])
def clean_working_place():
    global done_parcels
    done_parcels, parcel_plomb_numb = parcel_services.clean_working_place_service(done_parcels)

    object_name = parcel_plomb_numb
    comment = 'Отбор посылок: Завершено место'
    insert_user_action(object_name, comment)

    return render_template('parcel_info.html')


@app_svh.route('/search/clean_working_place_sql', methods=['POST'])
def clean_working_place_sql():
    parcel_plomb_numb = request.args.get('parcel_plomb_numb')
    print(parcel_plomb_numb)
    df_refuses = parcel_services.clean_working_place_sql_service(parcel_plomb_numb)
    if not df_refuses.empty:
        object_name = parcel_plomb_numb
        comment = f'Отбор посылок: Завершено место {parcel_plomb_numb}'
        insert_user_action(object_name, comment)
        flash(f'Завершено место {parcel_plomb_numb}', category='success')
        return render_template('parcel_search.html')
    else:
        flash(f'Ошибка завершения места (Проверьте user_id)!', category='error')
        return render_template('parcel_search.html')


parcel_plomb_numb = None


@app_svh.route('/search/manifest', methods=['GET'])
def make_manifest():
    try:
        global parcel_plomb_numb
        return render_template('plomb_search_manifest.html', search=parcel_plomb_numb)
    except Exception as e:
        return {'message': str(e)}, 400


df_plomb_to_manifest = pd.DataFrame()
df_plomb_to_manifest_total = pd.DataFrame()
df_to_manifest_HTML = pd.DataFrame()
df_refuse_plombs = pd.DataFrame()


@app_svh.route('/search/plomb_to_manifest', methods=['POST', 'GET'])
def plomb_to_manifest():
    global df_plomb_to_manifest
    global df_plomb_to_manifest_total
    global parcel_plomb_numb
    global df_to_manifest_HTML
    global df_refuse_plombs
    parcel_plomb_numb = request.form['parcel_plomb_numb']
    print(parcel_plomb_numb)
    if df_plomb_to_manifest.empty:
        df_plomb_to_manifest = pd.DataFrame({'parcel_plomb_numb': [parcel_plomb_numb]})
    try:
        con = sl.connect('BAZA.db')
        with con:
            df_plomb_to_manifest = pd.read_sql(f"SELECT * FROM baza where parcel_plomb_numb = '{parcel_plomb_numb}'",
                                               con)
            logger.warning(df_plomb_to_manifest['custom_status_short'])
            # if that is parcel:
            if df_plomb_to_manifest.empty:
                df_plomb_to_manifest = pd.read_sql(f"SELECT * FROM baza where parcel_numb = '{parcel_plomb_numb}'", con)
                df_all_pallet_plombs = df_plomb_to_manifest.drop_duplicates(subset=['parcel_plomb_numb', 'parcel_numb'],
                                                                            keep='first')
                df_all_pallet_plombs['Тип'] = "Посылка-место"
                df_plomb_to_manifest_refuse_parcel = df_all_pallet_plombs['parcel_numb'].values[0]
                if not df_plomb_to_manifest.loc[df_plomb_to_manifest['custom_status_short'] == 'ИЗЪЯТИЕ'].empty:
                    flash(f'ПОСЫЛКА {df_plomb_to_manifest_refuse_parcel} СО СТАТУСОМ ИЗЪЯТИЕ (НЕ ВЫПУЩЕНА!!!)',
                          category='error')
                    winsound.PlaySound('Snd\imanie izyyatie.wav', winsound.SND_FILENAME)
            # if that is plomb:
            else:
                df_plomb_to_manifest = df_plomb_to_manifest.drop_duplicates(subset='parcel_plomb_numb', keep='first')
                df_plomb_to_manifest['quont_plomb'] = None
                df_plomb_to_manifest['Тип'] = "Пломба"
                df_plomb_to_manifest['parcel_numb'] = ''
                pallet = df_plomb_to_manifest['pallet'].values[0]
                df_all_pallet_plombs = pd.read_sql(f"SELECT * FROM baza where pallet = '{pallet}'", con)
                writer = pd.ExcelWriter('df_all_pallet_plombs.xlsx', engine='xlsxwriter')
                df_all_pallet_plombs.to_excel(writer, sheet_name='Sheet1', index=False)
                writer.save()
                df_all_pallet_plombs_refuse = df_all_pallet_plombs.loc[
                    df_all_pallet_plombs['custom_status_short'] == 'ИЗЪЯТИЕ']
                if not df_all_pallet_plombs_refuse.empty:
                    df_all_pallet_plombs_refuse.drop_duplicates(subset='parcel_plomb_numb', keep='first')
                    df_all_pallet_plombs_refuse = df_all_pallet_plombs_refuse[
                        'parcel_plomb_numb'].drop_duplicates().to_string(index=False)
                    flash(f'НА ПАЛЛЕТЕ {pallet} ЕСТЬ ОТКАЗНЫЕ (НЕ ОТРАБОТАННЫЕ) ПЛОМБЫ: {df_all_pallet_plombs_refuse}',
                          category='error')
                    winsound.PlaySound('Snd\imanie izyyatie.wav', winsound.SND_FILENAME)
                df_all_pallet_plombs = df_all_pallet_plombs.drop_duplicates(subset='parcel_plomb_numb', keep='first')
                df_all_pallet_plombs['Тип'] = "Пломба"
                df_all_pallet_plombs['parcel_numb'] = ''
            df_all_pallet_plombs['quont_plomb'] = len(df_all_pallet_plombs)

            if df_all_pallet_plombs.empty:
                df_plomb_to_manifest_total = df_plomb_to_manifest_total.append(df_plomb_to_manifest)
                object_name = parcel_plomb_numb
                comment = 'Манифест огрузки: Пломба добавленна для отгрузки'
                insert_user_action(object_name, comment)
            else:
                df_plomb_to_manifest_total = df_plomb_to_manifest_total.append(df_all_pallet_plombs)
                object_name = pallet
                comment = 'Манифест огрузки: Паллет добавлен для отгрузки'
                insert_user_action(object_name, comment)

            df_plomb_to_manifest_total = df_plomb_to_manifest_total.drop_duplicates(
                subset=['parcel_plomb_numb', 'parcel_numb'], keep='first')
            df_plomb_to_manifest_total.index = df_plomb_to_manifest_total.index + 1  # shifting index
            #df_plomb_to_manifest_total.sort_index(inplace=True)
            df_plomb_to_manifest_total['№1'] = np.arange(len(df_plomb_to_manifest_total))[::-1] + 1
            #df_plomb_to_manifest_total['№1'] = np.arange(len(df_plomb_to_manifest_total))[::+1] + 1

            df_to_manifest_HTML = df_plomb_to_manifest_total[
                ['№1', 'parcel_plomb_numb', 'pallet', 'party_numb', 'quont_plomb', 'Тип', 'parcel_numb']]

            df_to_manifest_HTML = df_to_manifest_HTML.rename(columns={'№1': '№', 'parcel_plomb_numb': 'Пломба',
                                                                      'pallet': 'Паллет', 'party_numb': 'Партия',
                                                                      'quont_plomb': 'кол. пломб',
                                                                      'parcel_numb': 'Трек'})
            df_to_manifest_HTML.fillna("", inplace=True)

        object_name = parcel_plomb_numb
        comment = 'Манифест огрузки: Пломба добавленна для отгрузки'
        insert_user_action(object_name, comment)
    except Exception as e:
        flash(f'Пломба не найдена! {e}', category='error')
        winsound.PlaySound('Snd\Snd_NoPlomb.wav', winsound.SND_FILENAME)
    return render_template('plomb_info_manifest.html',
                           tables=[style + df_to_manifest_HTML.to_html(classes='mystyle', index=False)],
                           titles=['n/a', 'Добавленные'],
                           parcel_plomb_numb=parcel_plomb_numb)


@app_svh.route('/search/clean_working_place_manifest', methods=['POST'])
def clean_working_place_manifest():
    global df_plomb_to_manifest_total
    global df_to_manifest_HTML
    object_name = None
    comment = 'Манифест огрузки: Таблица отгрузки очищена'
    insert_user_action(object_name, comment)
    df_plomb_to_manifest_total = pd.DataFrame()
    flash('Манифест огрузки: Таблица отгрузки очищена', category='success')
    return render_template('plomb_info_manifest.html')


@app_svh.route('/save_manifest<string:partner>', methods=['POST', 'GET'])
def save_manifest(partner):
    global df_plomb_to_manifest_total
    con = sl.connect('BAZA.db', timeout=30)
    df_manifest_total_refuses = df_plomb_to_manifest_total.loc[
        df_plomb_to_manifest_total['custom_status_short'] == 'ИЗЪЯТИЕ']
    df_manifest_total_refuses_parcels = df_manifest_total_refuses.drop_duplicates(subset='parcel_numb')
    df_manifest_total_refuses_parcels = df_manifest_total_refuses_parcels['parcel_numb'].to_list()
    if df_manifest_total_refuses.empty:
        df_parcelc_to_manifest_total = df_plomb_to_manifest_total.loc[
            df_plomb_to_manifest_total['Тип'] == 'Посылка-место']
        df_plomb_to_manifest_total = df_plomb_to_manifest_total.loc[df_plomb_to_manifest_total['Тип'] == 'Пломба']
        with con:
            df_plombs_to_sql = df_plomb_to_manifest_total['parcel_plomb_numb']
            df_plombs_to_sql.to_sql('df_plomb_to_manifest', con=con, if_exists='replace')
        with con:
            query = """Select * FROM baza
                    JOIN df_plomb_to_manifest ON df_plomb_to_manifest.parcel_plomb_numb = baza.parcel_plomb_numb"""

            df_manifest_total = pd.read_sql(query, con)
            i = 0
            ploms_to_manifest = df_plomb_to_manifest_total['parcel_plomb_numb']

            for parcel_numb in df_parcelc_to_manifest_total['parcel_numb']:
                df_manifest = pd.read_sql(f"SELECT * FROM baza where parcel_numb = '{parcel_numb}'", con)
                con.execute("Update baza set VH_status = 'ОТГРУЖЕН' where parcel_numb = ?", (parcel_numb,))

                df_manifest_total = df_manifest_total.append(df_manifest)
            logger.warning(df_manifest_total)
        df_manifest_total = df_manifest_total.loc[:, ~df_manifest_total.columns.duplicated()].copy()
        writer = pd.ExcelWriter('test.xlsx', engine='xlsxwriter')
        df_manifest_total.to_excel(writer, sheet_name='Sheet1', index=False)
        writer.save()
        if partner == 'CEL':
            manifest_services.manifest_to_xls(df_manifest_total)
        elif partner == 'GBS':
            manifest_services.manifest_to_xls_GBS(df_manifest_total)
        print(ploms_to_manifest)
        for parcel_plomb_numb in ploms_to_manifest:
            print(parcel_plomb_numb)
            with con:
                i += 1
                con.execute(
                    "Update baza set VH_status = 'ОТГРУЖЕН' where parcel_plomb_numb = ? and custom_status_short = ?",
                    (parcel_plomb_numb, 'ВЫПУСК'))
                print(f'{i} - {parcel_plomb_numb} updated')
        logger.warning(df_manifest_total)
    else:
        flash(f'{df_manifest_total_refuses_parcels} СО СТАТУСОМ ИЗЪЯТИЕ (НЕ ВЫПУЩЕНА!!!)', category='error')
        winsound.PlaySound('Snd\imanie izyyatie.wav', winsound.SND_FILENAME)
    return render_template('plomb_info_manifest.html')


@app_svh.route('/party_to_manifest', methods=['POST', 'GET'])
def party_to_manifest():
    con = sl.connect('BAZA.db')
    with con:
        df_parties = pd.read_sql(f"SELECT * FROM baza", con).drop_duplicates(subset='party_numb', keep='first')
        df_parties = df_parties['party_numb'].to_list()
    return render_template('party_to_manifest.html', df_parties=df_parties)


@app_svh.route('/party_to_manifest_button', methods=['POST', 'GET'])
def party_to_manifest_button():
    party_numb = request.form['party_numb']
    con = sl.connect('BAZA.db')
    with con:
        df_manifest_total = pd.read_sql(f"SELECT * FROM baza where party_numb = '{party_numb}' "
                                        f"and custom_status_short = 'ВЫПУСК'", con)
        con.execute("Update baza set VH_status = 'ОТГРУЖЕН' where party_numb = ? AND custom_status_short = ?",
                    (party_numb, 'ВЫПУСК'))
    df_manifest_total = df_manifest_total.sort_values(by='parcel_plomb_numb', ascending=False)
    object_name = party_numb
    comment = 'Манифест огрузки: Партия отгружена'
    insert_user_action(object_name, comment)
    manifest_services.manifest_to_xls(df_manifest_total)
    return render_template('plomb_search_manifest.html')


df_changed_plomb = pd.DataFrame()


@app_svh.route('/search1/old_plomb/add', methods=['GET', 'POST'])
def old_plomb_add():
    now_time = datetime.datetime.now().strftime("%d.%m.%Y %H:%M")
    global df_changed_plomb
    try:
        if request.method == 'POST':
            old_plomb_numb = request.form['old_plomb_numb']
            new_plomb_numb = request.form['new_plomb_numb']
            con = sl.connect('BAZA.db')
            with con:
                df = pd.read_sql(f"SELECT * FROM baza where parcel_plomb_numb = '{old_plomb_numb}' ", con)
                if df.empty:
                    flash(f'Пломба {old_plomb_numb} не найдена!', category='error')
                    winsound.PlaySound('Snd\Snd_NoPlomb.wav', winsound.SND_FILENAME)
                else:

                    cursor = con.cursor()
                    query = "Update baza set parcel_plomb_numb = ? where parcel_plomb_numb = ?"
                    data = (new_plomb_numb, old_plomb_numb)
                    query2 = "Update plombs_with_pullings set parcel_plomb_numb = ? where parcel_plomb_numb = ?"
                    logger.warning(data)
                    cursor.execute(query, data)
                    cursor.execute(query2, data)
                    con.commit()
                    cursor.close()
                    flash(f'Пломба {old_plomb_numb} обновлена!', category='success')
                    winsound.PlaySound('Snd\plomba_obnovlena.wav', winsound.SND_FILENAME)
                    df_changed_plomb_to_append = pd.DataFrame({'старая': [old_plomb_numb], 'новая': [new_plomb_numb]})
                    df_changed_plomb = df_changed_plomb.append(df_changed_plomb_to_append)
                    df_changed_plomb['№'] = np.arange(len(df_changed_plomb))[::+1] + 1
                    df_changed_plomb = df_changed_plomb[['№', 'старая', 'новая']]
                    df_changed_plomb = df_changed_plomb.iloc[::-1]
                    logger_change_plob.info(
                        f'Время: {now_time} Старая пломба: {old_plomb_numb} Новая пломба: {new_plomb_numb}')
                    object_name = old_plomb_numb
                    comment = f'Новая пломба: Изменена на {new_plomb_numb}'
                    insert_user_action(object_name, comment)
        return render_template('old_plomb.html',
                               tables=[style + df_changed_plomb.to_html(classes='mystyle', index=False,
                                                                        float_format='{:2,.2f}'.format)],
                               titles=['na', 'Изменены:'])

    except Exception as e:
        logger.warning(f'error {e}')
        return {'message': str(e)}, 400


@app_svh.route('/new_place', methods=['POST', 'GET'])
def new_place():
    try:
        parcel_numb = request.args.get('parcel_numb')
        object_name = parcel_numb
        comment = f'Сформировать место: Посылка отмечена для формирования нового места'
        insert_user_action(object_name, comment)
        return render_template('parcel_search_new_place.html', search=parcel_numb)
    except Exception as e:
        return {'message': str(e)}, 400


done_parcels_np = pd.DataFrame()


def style_df(df_parc_events_np):
    done_parcels_styl_np = done_parcels_np.reset_index()
    done_parcels_styl_np = done_parcels_styl_np.drop('index', axis=1)
    done_parcels_styl_np = done_parcels_styl_np[
        ['№', 'Трек-номер', 'Статус', 'Пломба', 'ВХ']].drop_duplicates(subset=['Трек-номер'], keep='first')
    trigger_color1 = df_parc_events_np['Статус']

    # trigger_color2 = df_parc_events_np['ВХ']
    # def highlight_GREEN(df_parc_events):
    #    return ['background-color: #7CFC00' if 'На ВХ' in str(i) else '' for i in df_parc_events]
    def highlight_RED(df_parc_events_np):
        return ['background-color: #FF0000' if 'ИЗЪЯТИЕ' in str(i) else '' for i in df_parc_events_np]
    if 'ИЗЪЯТИЕ' in done_parcels_styl_np['Статус'].values:
        done_parcels_styl_np = done_parcels_styl_np.style.apply(highlight_RED).hide_index()
    return done_parcels_styl_np

@app_svh.route('/save_new_place', methods=['POST', 'GET'])
def making_new_place():
    global done_parcels_np
    global parcel_numb_np
    global df_parc_events_np
    global done_parcels_styl_np
    parcel_numb_np = request.form['parcel_numb']
    con_vec = sl.connect('VECTORS.db')
    try:
        vector = con_vec.execute(
            f"SELECT vector FROM VECTORS where parcel_numb = '{parcel_numb_np}'").fetchone()[
            0]
        flash(f'{vector}!', category='success')
    except Exception as e:
        vector = None
        print(e)

    try:
        con = sl.connect('BAZA.db')
        with con:
            df_parc_events_np = pd.read_sql(f"SELECT * FROM baza where parcel_numb = '{parcel_numb_np}'", con)
            if df_parc_events_np['custom_status_short'].values[0] == 'ИЗЪЯТИЕ':

                if 'Требуется' in df_parc_events_np['custom_status'].values[0] or 'уплат' in \
                        df_parc_events_np['refuse_reason'].values[0] or 'Не уплачены' in \
                        df_parc_events_np['refuse_reason'].values[0]:
                    pay_trigger = 'ПЛАТНАЯ'
                else:
                    pay_trigger = ''
                flash(f'ИЗЪЯТИЕ {pay_trigger}', category='error')
            df_parc_events_np['№1'] = np.arange(len(df_parc_events_np))[::+1] + 1
            df_parc_events_np = df_parc_events_np[
                ['parcel_numb', 'custom_status_short', 'parcel_plomb_numb', 'VH_status']]
            df_parc_events_np = df_parc_events_np.rename(
                columns={'parcel_numb': 'Трек-номер', 'custom_status_short': 'Статус',
                         'parcel_plomb_numb': 'Пломба', 'VH_status': 'ВХ'})
            status = df_parc_events_np['Статус'][0]
            if df_parc_events_np.empty:
                flash(f'Посылка не найдена!', category='error')
                winsound.PlaySound('Snd\Snd_Parcel_Not_Found.wav', winsound.SND_FILENAME)
            else:
                pass
            done_parcels_np = done_parcels_np.append(df_parc_events_np).drop_duplicates(subset=['Трек-номер'],
                                                                                        keep='last')
            done_parcels_np.index = done_parcels_np.index + 1  # shifting index
            done_parcels_np.sort_index(inplace=True)
            done_parcels_np['№'] = np.arange(len(done_parcels_np))[::-1] + 1
            done_parcels_styl_np = style_df(df_parc_events_np)
            object_name = parcel_numb_np
            comment = f'Сформировать место: Посылка отмечена для формирования нового места'
            insert_user_action(object_name, comment)
    except Exception as e:
        flash(f'Посылка не найдена!', category='error')
        winsound.PlaySound('Snd\Snd_Parcel_Not_Found.wav', winsound.SND_FILENAME)
        print(e)
        return render_template('parcel_info_new_place.html')
        # return {'message': str(e)}, 400
        pass
    return render_template('parcel_info_new_place.html',
                           tables=[done_parcels_styl_np.to_html(classes='mystyle', index=False,
                                                                float_format='{:2,.2f}'.format)],
                           titles=['na', '\n\nОтработанные:'],
                           parcel_numb=parcel_numb_np, vector=vector)


@app_svh.route('/delete_last_parcel', methods=['POST', 'GET'])
def delete_last_parcel():
    global done_parcels_np
    global df_parc_events_np
    global done_parcels_styl_np
    logger.warning(done_parcels_np)
    done_parcels_np = done_parcels_np[1:]
    done_parcels_styl_np = style_df(df_parc_events_np)
    logger.warning(done_parcels_np)
    return render_template('parcel_info_new_place.html',
                           tables=[df_parc_events_np.to_html(classes='mystyle', index=False,
                                                             float_format='{:2,.2f}'.format),
                                   style + done_parcels_styl_np.to_html(classes='mystyle',
                                                                        index=False,
                                                                        float_format='{:2,.2f}'.format)],
                           titles=['na', 'Посылка:', '\n\nОтработанные:'],
                           parcel_numb=parcel_numb_np)


@app_svh.route('/save_new_place/place_numb', methods=['POST', 'GET'])
def make_place_numb():
    global done_parcels_np
    plomb_toreplace_new = request.form['plomb_toreplace_new']
    con = sl.connect('BAZA.db')
    # открываем базу
    with con:
        for parcel_numb in done_parcels_np['Трек-номер']:
            # plomb_toreplace = done_parcels_np.loc[done_parcels_np['Трек-номер'] == parcel_numb]['Пломба'].values[0]
            # logger.warning(plomb_toreplace)
            con.execute(
                f"Update baza set parcel_plomb_numb = '{plomb_toreplace_new}' where parcel_numb = '{parcel_numb}'")
    df_new_place = pd.read_sql(f"Select * from baza where parcel_plomb_numb = '{plomb_toreplace_new}'", con)
    writer = pd.ExcelWriter(f'{addition_folder}Место {plomb_toreplace_new}.xlsx', engine='xlsxwriter')
    df_new_place.to_excel(writer, sheet_name='Sheet1', index=False)
    for column in df_new_place:
        column_width = max(df_new_place[column].astype(str).map(len).max(), len(column))
        col_idx = df_new_place.columns.get_loc(column)
        writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_width)
        writer.sheets['Sheet1'].set_column(0, 3, 10)
        writer.sheets['Sheet1'].set_column(1, 3, 20)
        writer.sheets['Sheet1'].set_column(2, 3, 20)
        writer.sheets['Sheet1'].set_column(3, 3, 20)
        writer.sheets['Sheet1'].set_column(4, 3, 30)
        writer.sheets['Sheet1'].set_column(5, 3, 20)
    object_name = plomb_toreplace_new
    comment = f'Сформировать место: Место сформированно'
    insert_user_action(object_name, comment)
    writer.save()
    con.commit()
    con.close()
    done_parcels_np = pd.DataFrame()
    flash(f'Место сформировано', category='success')
    winsound.PlaySound('Snd\mesto_sformirovano.wav', winsound.SND_FILENAME)
    return render_template('parcel_info_new_place.html')


@app_svh.route('/making_new_place_sql', methods=['POST', 'GET'])
def making_new_place_sql():
    try:
        parcel_numb, vector, audiofile, df_user_work = add_to_place_sql_service()
        print(parcel_numb)
    except:
        flash(f'Посылка не найдена!', category='error')
        print(str(traceback.format_exc()))
        parcel_numb = None
        vector = None
        audiofile = 'Snd_CancelIssue.wav'
        df_user_work = pd.DataFrame()
    return render_template('parcel_info_new_place_sql.html', parcel_numb=parcel_numb, audiofile=audiofile,
                           vector=vector, tables=[style + df_user_work.to_html(classes='mystyle', index=False,
                                                                float_format='{:2,.2f}'.format)],
                           titles=['na', '\n\nОтработанные:'])


@app_svh.route('/delete_last_parcel_place_sql', methods=['POST', 'GET'])
def delete_last_place_sql():
    user_name, user_id = get_user_name()
    con = sl.connect("BAZA.db")
    with con:
        df_add_to_place = pd.read_sql(
            f"SELECT * FROM add_to_place where user_id = '{user_id}'",
            con)
        last_id = df_add_to_place['ID'].max()
        con.execute(
            f"DELETE FROM add_to_place where ID = '{last_id}'")
        df_user_work = df_add_to_place.fillna('').sort_values(by='ID', ascending=False)

        df_user_work['№'] = np.arange(len(df_user_work))[::-1] + 1
        df_user_work = df_user_work[['№', 'parcel_numb', 'parcel_plomb_numb', 'custom_status_short', 'user_id']]
        print(last_id)
        object_name = last_id
        comment = 'Сформировать место: удалена последняя запись'
        insert_user_action(object_name, comment)
    print('ok')
    return render_template('parcel_info_new_place_sql.html',
                           tables=[style + df_user_work.to_html(classes='mystyle', index=False,
                                                                float_format='{:2,.2f}'.format)],
                           titles=['na', '\n\nОтработанные:'])


@app_svh.route('/add_to_place_button', methods=['POST', 'GET'])
def add_to_place_button():
    user_name, user_id = get_user_name()
    parcel_plomb_numb = request.form['parcel_plomb_numb']
    print(parcel_plomb_numb)
    con = sl.connect("BAZA.db")
    with con:
        df_add_to_place = pd.read_sql(
                    f"SELECT * FROM add_to_place where user_id = '{user_id}'",
                    con)
        for parcel_numb in df_add_to_place['parcel_numb']:
            con.execute(
                f"Update baza set parcel_plomb_numb = '{parcel_plomb_numb}',"
                f"VH_status = 'Готово к отгрузке', "
                f"zone = '' where parcel_numb = '{parcel_numb}'")

            print(parcel_numb)
        object_name = parcel_plomb_numb
        comment = 'Сформировать место: сформированно новое место'
        insert_user_action(object_name, comment)
    qnt_df_add_to_place = len(df_add_to_place)
    with con:
        con.execute(
            f"DELETE FROM add_to_place where user_id = '{user_id}'")
    flash(f'Место {parcel_plomb_numb} сформированно, кол-во посылок: {qnt_df_add_to_place}', category='success')
    #print('ok')
    return render_template('parcel_info_new_place_sql.html')


@app_svh.route('/create_zone', methods=['POST', 'GET'])
def add_to_zone():
    try:
        parcel_numb, vector, audiofile, df_user_work = add_to_zone_service()
        print(parcel_numb)
    except:
        flash(f'Посылка не найдена!', category='error')
        print(str(traceback.format_exc()))
        parcel_numb = None
        vector = None
        audiofile = 'Snd_CancelIssue.wav'
        df_user_work = pd.DataFrame()
    return render_template('parcel_info_new_zone.html', parcel_numb=parcel_numb, audiofile=audiofile,
                           vector=vector, tables=[style + df_user_work.to_html(classes='mystyle', index=False,
                                                                float_format='{:2,.2f}'.format)],
                           titles=['na', '\n\nОтработанные:'])


@app_svh.route('/delet_last_parcel_zone', methods=['POST', 'GET'])
def delet_last_parcel_zone():
    user_name, user_id = get_user_name()
    con = sl.connect("BAZA.db")
    with con:
        df_add_to_zone = pd.read_sql(
            f"SELECT * FROM add_to_zone where user_id = '{user_id}'",
            con)
        last_id = df_add_to_zone['ID'].max()
        con.execute(
            f"DELETE FROM add_to_zone where ID = '{last_id}'")
        df_user_work = df_add_to_zone.fillna('').sort_values(by='ID', ascending=False)

        df_user_work['№'] = np.arange(len(df_user_work))[::-1] + 1
        df_user_work = df_user_work[['№', 'parcel_numb', 'zone', 'user_id']]
        print(last_id)
        object_name = last_id
        comment = 'Сформировать зону хранения: удалена последняя запись'
        insert_user_action(object_name, comment)
    print('ok')
    return render_template('parcel_info_new_zone.html',
                           tables=[style + df_user_work.to_html(classes='mystyle', index=False,
                                                                float_format='{:2,.2f}'.format)],
                           titles=['na', '\n\nОтработанные:'])


@app_svh.route('/add_to_zone_button', methods=['POST', 'GET'])
def add_to_zone_button():
    user_name, user_id = get_user_name()
    zone = request.form['zone']
    print(zone)
    con = sl.connect("BAZA.db")
    with con:
        df_add_to_zone = pd.read_sql(
                    f"SELECT * FROM add_to_zone where user_id = '{user_id}'",
                    con)
        for parcel_numb in df_add_to_zone['parcel_numb']:
            con.execute(
                f"Update baza set zone = '{zone}' where parcel_numb = '{parcel_numb}'")
            print(parcel_numb)
        qnt_df_add_to_zone = len(df_add_to_zone)
        con.execute(
            f"DELETE FROM add_to_zone where user_id = '{user_id}'")
        object_name = zone
        comment = f'Посылки размещенны в зоне {zone} в кол-ве: {qnt_df_add_to_zone}'
        insert_user_action(object_name, comment)
    flash(f'Посылки размещенны в зоне {zone} в кол-ве: {qnt_df_add_to_zone}', category='success')
    #print('ok')
    return render_template('parcel_info_new_zone.html')


@app_svh.route('/modal')
def modal():
    return render_template('modal.html')

@app_svh.route('/search', methods=['GET'])
def parc_searh():
    return render_template('parc_search.html')


@app_svh.route('/get_info', methods=['POST', 'GET'])
def get_parcel_info_list():
    global df_all_parcels
    parcel_numbs = request.form['parcel_numbs'].replace(' ', ',')
    parcels_list = parcel_numbs.split(",")
    df_all_parcels = pd.DataFrame()
    for parcel_numb in parcels_list:
        try:
            con = sl.connect('BAZA.db')
            with con:
                df_parc_events = pd.read_sql(
                    f"SELECT * FROM baza where parcel_numb = '{parcel_numb}' ORDER BY ID", con)
                print(df_parc_events)
                df_all_parcels = df_all_parcels.append(df_parc_events)
        except Exception as e:
            logger.warning(f'parcel {parcel_numb}  - read action faild with error: {e}')
            return {'message': str(e)}, 400
    return render_template('parc_info.html', tables=[df_all_parcels.to_html(classes='mystyle', index=False)],
                           titles=['na', 'ALL'])


@app_svh.route('/info/parc_info_to_xl', methods=['GET', 'POST'])
def parc_info_to_xl():
    global df_all_parcels
    print(df_all_parcels)
    now_time = datetime.datetime.now().strftime("%d.%m.%Y %H-%M")
    writer = pd.ExcelWriter(f'{download_folder}Выгрузка по посылкам от {now_time}.xlsx', engine='xlsxwriter')
    df_all_parcels.to_excel(writer, sheet_name='Sheet1', index=False)
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
    flash(f'Инфо по посылкам выгружена в excel!', category='success')
    object_name = ''
    comment = f'Инфо по ппосылкам выгружена в excel!'
    insert_user_action(object_name, comment)
    print('writer ok')
    return render_template('parc_info.html', tables=[df_all_parcels.to_html(classes='mystyle', index=False)],
                           titles=['na', 'ALL'])


@app_svh.route('/api/get_decisions', methods=['GET', 'POST'])
def load_decision_api():
    server_request_events()
    check_and_backup()
    return render_template('index2.html')


def chunks(lst, n):
    """Yield successive n-sized chunks from lst."""
    for i in range(0, len(lst), n):
        yield lst[i:i + n]


def tochina_prepare(parcel_numb, custom_status, refuse_reason, Event_date_chin, event_code):
    track_codes = {
                "ERT": "HCPR",
                "RIC": "SHND",
                "PCD": "SHND",
                "HBA45": "HCCR",
                "HBA44": "HCCR",
                "HBA41": "HCGR",
                "ASF5": "HCFK",
                "CR": "RC",
                "CR2": "RC",
                "CR3": "RC",

                }
    track_chinees = {'HCPR': '延期放行', 'SHND': '需缴纳关税', 'HCCR': '护照无效', 'HCSM': '产品个人使用说明',
                     'HCCP': '需提供产品说明书', 'HCGR': 'B2B（商品',
                     'HCGS': 'B2B（数量', 'HCFK': '需提供付款凭证', 'HCRU': '需提供网址链接'}

    try:
        track_code = track_codes[event_code]
        chines_custom_status = track_chinees[track_code]
    except:
        track_code = ''
        chines_custom_status = ''
    if event_code == 'IDOK':
        chines_custom_status = '护照数据已被海关经纪人接受'
    try:
        decision_date = datetime.datetime.strftime(Event_date_chin, "%Y/%#m/%#d %H:%M")
        if refuse_reason == 'nan':
            refuse_reason = ''
        else:
            for key, item in track_codes.items():
                if key in str.lower(refuse_reason):
                    track_code = item
                    break

        if 'выпуск товаров' in str.lower(custom_status):
            track_code = 'RC'
        elif 'продление' in str.lower(custom_status):
            track_code = 'HCPR'


        for key, item in track_chinees.items():
            if track_code != 'RC' and track_code == key:
                refuse_reason = item + refuse_reason
                break
            if track_code == 'RC':
                custom_status = custom_status + '放行'
                break
        if refuse_reason != '':
            refuse_reason = refuse_reason + '. '
        print(track_code)
        print(f"{custom_status + refuse_reason}")
        data = {"PostingNumber": f"{parcel_numb}", "TrackingNumber": f"{parcel_numb}",
                "Data": [{"track_code": f"{track_code}", "datetime": f"{decision_date}", "location": "Россия",
                          "description": f"{chines_custom_status} {custom_status + refuse_reason}"}]}
        return data
    except Exception as e:
        logger_GBS_statuses.info(f'insert_event_API action faled: {parcel_numb}: {e}')


def send_to_china(data):
    try:
        print(data)
        data_str = str(data).replace("'", '"').replace(", ", ",")
        m = hashlib.md5()
        m.update(data_str.encode('utf-8'))
        result = base64.urlsafe_b64encode(m.hexdigest().encode('utf-8')).decode(
            'utf-8')  # b64encode(m.hexdigest().encode('utf-8'))
        url = ("http://hccd.rtb56.com/webservice/edi/TrackService.ashx?code=ADDCUSTOMSCLEARANCETRACK"
               + f'&data={data_str}' + f'&sign={str(result)}')
        response = requests.get(url)
        logger_GBS_statuses.info(f'insert_event_API action: {response.text}')
        print(response.text)
        print('ok')
    except Exception as e:
        logger_GBS_statuses.info(f'send_to_china faled: {data}: {e}')
        pass


delta = datetime.timedelta(hours=-10, minutes=0)


def creating_pay_info_GBS(parcel_numb, Event_date_chin):

    try:
        expired_date = Event_date_chin + delta
        expired_date = expired_date.strftime("%Y-%m-%d")
        print(expired_date)
        url = "http://hccd.rtb56.com/webservice/Ozon/OzonSavePayTaxData.ashx"
        data = [
            {
                "posting_number": '',
                "tracking_number": parcel_numb,
                "pay_tax_end_time": expired_date,
                "pay_tax_link": "https://gbs-broker.alta.ru/",
                "tax_amount": '',
                "is_paid": "N"
            }
        ]
        response = requests.post(url=url, json=data)
        print(response.text)
        logger_GBS_statuses.info(f'{now} creating_pay_info OK: {parcel_numb}')
        return "OK"
    except Exception:
        logger_GBS_statuses.info(f'{now} creating_pay_info faled: {str(traceback.format_exc())}')
        return (str(traceback.format_exc()))


def payresult_GBS(parcel_numb):
    try:
        url = 'http://hccd.rtb56.com/webservice/Ozon/OzonUpdatePayTaxData.ashx'
        data = [
            {
                "tracking_number": f"{parcel_numb}",
                "is_paid": "Y"
            }
        ]
        response = requests.post(url=url, json=data)
        logger_GBS_statuses.info(f'{now} pay result OK: {parcel_numb} {response.text}:')
        return "OK"
    except Exception:
        logger_GBS_statuses.info(f'{now} pay_result faled: {str(traceback.format_exc())}')
        return (str(traceback.format_exc()))


def GBS_request_events():
    con = sl.connect('BAZA.db')
    len_id = con.execute('SELECT max(id) from baza').fetchone()[0]
    print(len_id)
    id_for_job = len_id - 5000000
    print(id_for_job)
    df = pd.read_sql(f"Select parcel_numb from baza "
                     f"where party_numb LIKE '%URC%' "
                     f"AND ID > {id_for_job} "  # where ID > (len(ID) - 200 000)
                     f"AND custom_status_short = 'ИЗЪЯТИЕ' "
                     f"AND custom_status != 'Return in process'", con).drop_duplicates(subset='parcel_numb')
    print(df)
    list_chanks = list(chunks(df['parcel_numb'], 25))
    #print(list_chanks)
    n = 0
    for chank in list_chanks:
        print(chank)
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
        try:
            response = requests.post(url=url, json=body,  # http://164.132.182.145:5001
                                 headers=headers)
        except requests.exceptions.ConnectionError:
            time.sleep(3)
            response = requests.post(url=url, json=body,  # http://164.132.182.145:5001
                                     headers=headers)

        print(response)
        json_events = response.json()
        print(json_events)
        code_mask = {"NW": "Получена информация об отправлении",
                        "NWC": "Отменен",
                        "CSW": "Отправлено на склад 1-ой мили.",
                        "WA": "Принято на складе 1-ой мили.",
                        "WD": "Убыло со склада 1-й мили.",
                        "NWCFG": "Посылка отклонена перевозчиком, так как содержит запрещённые товары",
                        "NWCFM": "Отменен на первой миле",
                        "NWCHM": "Посылка отклонена перевозчиком, так как содержит опасные материалы",
                        "ZX": "Обработано на складе перевозчика.",
                        "ZC": "Готово к вылету в страну получателя.",
                        "RS": "Покинуло страну происхождения",
                        "AI": "Прибыло в страну назначения.",
                        "CF1": "Отправление не прилетело",
                        "CF3": "Товары изъяты таможенными органами",
                        "CNP5": "Сбой информационной системы таможенных органов",
                        "DS": "Задержка вылета.",
                        "CT": "Таможенный транзит",
                        "CI": "Прибыло на таможню.",
                        "CR": "Выпущено таможенным органом.",
                        "CR2": "Выпущено таможенным органом без платежей.",
                        "CR3": "Выпущено таможенным органом с платежом.",
                        "RSW": "Отправление готово к отгрузке со склада",
                        "CO": "Убыло из таможни.",
                        "PIB": "Товар поврежден (повреждение упаковки груза)",
                        "RPC": "Отказ от уплаты таможенных платежей",
                        "RPI": "Отказ от предоставления информации получателем",
                        "ASF": "Проблема. Запрет от таможни или секюрити.",
                        "ASF1": "Проблема. Товары не для личного пользования.",
                        "ASF2": "Проблема. Ссылка на товар не рабочая.",
                        "ASF3": "Проблема. Стоимость товара некорректна.",
                        "ASF4": "Проблема. Необходимо подтверждение паспортных данных для таможни.",
                        "ASF5": "Проблема. Запрос документов и сведений от таможенных органов.",
                        "HBA": "Проблема. Задержано таможней.",
                        "HBA41": "Отказ в выпуске. Не для личного пользования.",
                        "HBA42": "Отказ в выпуске. Не предоставлены документы.",
                        "HBA43": "Отказ в выпуске технического характера.",
                        "HBA44": "Отказ в выпуске. Некорректные ПД.",
                        "HBA45": "Отказ в выпуске. Недействительные ПД.",
                        "HBA46": "Отказ в выпуске. Иное.",
                        "ERT": "Продление срока выпуска.",
                        "OH": "Временное хранение",
                        "NOID": "Проблема. Нет паспортных данных в момент check in на таможне.",
                        "IDOK": "Паспортные данные собраны",
                        "DOCOK": "Документы для подачи в ТО представлены.",
                        "IDCE": "Время для сбора паспортных данных истекло",
                        "RIC": "Квитанция выставлена таможенным органом",
                        "RPR": "Квитанция оплачена получателем",
                        "PCD": "Отказ в выпуске(требуется уплата таможенных платежей)"}
        result = json_events['result']
        with con:
            parcel_list = []
            for parcel_slot in result:
                try:
                    parcel_numb = parcel_slot['HWBRefNumber']
                    event = parcel_slot['events'][0]
                    event_code = event['event_code']
                    try:
                        custom_status = code_mask[event_code]
                    except:
                        custom_status = event['event_text']
                    events_all = parcel_slot['events']
                    if 'clearance complete' in custom_status or 'Released by customs' in custom_status:
                        custom_status_short = 'ВЫПУСК'
                        decision_date = parcel_slot['events'][0]['event_time']
                        refuse_reason = parcel_slot['events'][0]['event_comment']
                    else:
                        for event in events_all:
                            if 'CR' in event['event_code'] or 'CR2' in event['event_code'] or 'CR3' in event[
                                'event_code']:
                                custom_status_short = 'ВЫПУСК'
                                decision_date = event['event_time']
                                refuse_reason = event['event_comment']
                                break
                            else:

                                custom_status_short = 'ИЗЪЯТИЕ'
                                decision_date = event['event_time']
                                refuse_reason = event['event_comment']
                                print(refuse_reason)

                        # custom_status_short = 'ИЗЪЯТИЕ'

                    con.execute(f"Update baza set "
                                f" custom_status = '{custom_status}',"
                                f" custom_status_short = '{custom_status_short}',"
                                f" decision_date = '{decision_date}',"
                                f" refuse_reason = '{refuse_reason}' where parcel_numb = '{parcel_numb}'")
                    print('updated')
                    Event_date_chin = datetime.datetime.strptime(decision_date, "%Y-%m-%dT%H:%M:%S")
                    data = tochina_prepare(parcel_numb, custom_status, refuse_reason, Event_date_chin, event_code)
                    send_to_china(data)

                    if event_code == 'RIK':
                        creating_pay_info_GBS(parcel_numb, Event_date_chin)
                    if event_code == 'RPR':
                        payresult_GBS(parcel_numb)
                    decision_date = decision_date.replace('T', ' ')
                    parcel_info = {"regnumber": '', "parcel_numb": parcel_numb,
                                   "Event": custom_status, "Event_comment": refuse_reason,
                                   "Event_date": decision_date}
                    parcel_list.append(parcel_info)
                except Exception as e:
                    logger.warning(f'parcel_slot: {parcel_slot} - ERROR: {e}')

        list_chunks_parc = list(chunks(parcel_list, 25))
        print(list_chunks_parc)
        i = 0
        for chunk_parc in list_chunks_parc:
            i += 1
            print(f'chunk {i}')
            response = requests.post('http://164.132.182.145:5000/api/add/new_event_chunks2', json=chunk_parc,
                                     headers={'accept': 'application/json'})
            print(response.text)

        n += 25
        print(n)


scheduler = BackgroundScheduler(daemon=True)
#Create the job

scheduler.add_job(func=GBS_request_events, trigger='interval', seconds=10) #trigger='cron', hour='22', minute='30'
scheduler.start()

if __name__ == '__main__':
    app_svh.secret_key = 'c9e779a3258b42338334daaed51bccf7'
    app_svh.config['SESSION_TYPE'] = 'filesystem'
    serve(app_svh, host='0.0.0.0', port=5000, threads=4)
