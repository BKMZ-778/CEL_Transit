from flask import Blueprint
from flask import render_template
from flask import flash
import pandas as pd
import sqlite3 as sl
import numpy as np
from SVH_BAZA_modules.services import (insert_user_action, addition_folder,)


bp_party_to_pallet = Blueprint('party_to_pallet', __name__, url_prefix='/party_to_pallet')

@bp_party_to_pallet.route('/party_info_allnotshipped_to_pallet/<string:row>', methods=['GET', 'POST'])
def party_info_allnotshipped_to_pallet(row):
    con = sl.connect('BAZA.db')
    try:
        with con:
            df_all_pallets = pd.read_sql(f"SELECT pallet FROM baza", con).drop_duplicates(subset='pallet', keep='first')

            try:
                df_all_pallets['pallet'] = df_all_pallets['pallet'].fillna(value=np.nan).fillna(0).astype(int)
                df_all_pallets = df_all_pallets.sort_values(by='pallet')
                last_pall_numb = int(df_all_pallets.values[-1].tolist()[0])
                new_pallet_numb = last_pall_numb + 1
            except:
                new_pallet_numb = 0

            con.execute(
                f"Update baza set pallet = '{new_pallet_numb}' where party_numb = '{row}' and VH_status is ?", (None,))
            df_new_pallet = pd.read_sql(f"Select * from baza where pallet = '{new_pallet_numb}'", con)
            writer = pd.ExcelWriter(f'{addition_folder}Паллет {new_pallet_numb}.xlsx', engine='xlsxwriter')
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
    except Exception as e:
        return {'message': str(e)}, 400

    flash(f'Паллет {new_pallet_numb} c не отгруженными из {row} сформирован!', category='success')
    object_name = new_pallet_numb
    comment = f'Паллет c не отгруженными из {row} сформирован!'
    insert_user_action(object_name, comment)
    return render_template('party_info.html', row=row
                           )


@bp_party_to_pallet.route('/party_info_vectors_to_pallet/<string:row>', methods=['GET', 'POST'])
def party_info_vectors_to_pallet(row):
    con = sl.connect('BAZA.db')
    print(row)
    try:
        df_party = pd.read_sql(f"SELECT * FROM baza where party_numb = '{row}'", con)
        df_party = df_party.loc[df_party['parcel_plomb_numb'] != '']
        con_vect = sl.connect('VECTORS.db')
        with con_vect:
            df_data = pd.read_sql(f"SELECT * FROM vectors where party_numb = '{row}'", con_vect)
            df_data = df_data.loc[df_data['parcel_plomb_numb'] != '']
            df_merge_vector_party = pd.merge(df_party, df_data,
                                             how='left',
                                             left_on='parcel_numb',
                                             right_on='parcel_numb').drop_duplicates(subset='parcel_plomb_numb_x')
            print(df_merge_vector_party)
            df_vectors = df_merge_vector_party.drop_duplicates(subset='vector')
            dict_of_pallets = {}
            for vector in df_vectors['vector']:
                with con:
                    df_all_pallets = pd.read_sql(f"SELECT pallet FROM baza", con).drop_duplicates(subset='pallet',
                                                                                                  keep='first')
                    try:
                        df_all_pallets['pallet'] = df_all_pallets['pallet'].fillna(value=np.nan).fillna(0).astype(int)
                        df_all_pallets = df_all_pallets.sort_values(by='pallet')
                        last_pall_numb = int(df_all_pallets.values[-1].tolist()[0])
                        new_pallet_numb = last_pall_numb + 1
                    except:
                        new_pallet_numb = 0
                    print(new_pallet_numb)
                    print(vector)
                    dict_of_pallets[vector] = new_pallet_numb
                    df_create_pallet_vector = df_merge_vector_party.loc[df_merge_vector_party['vector'] == vector]
                    for parcel_plomb_numb in df_create_pallet_vector['parcel_plomb_numb_x']:
                        con.execute(f"Update baza set pallet = '{new_pallet_numb}' where parcel_plomb_numb = '{parcel_plomb_numb}' and VH_status is ?", (None,))
                df_new_pallet = pd.read_sql(f"Select * from baza where pallet = '{new_pallet_numb}'", con)
                writer = pd.ExcelWriter(f'{addition_folder}Паллет {new_pallet_numb} - {vector}.xlsx', engine='xlsxwriter')
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
    except Exception as e:
        return {'message': str(e)}, 400

    flash(f'Паллеты {dict_of_pallets} c не отгруженными из {row} сформирован!', category='success')
    object_name = new_pallet_numb
    comment = f'Паллет c не отгруженными из {row} сформирован!'
    insert_user_action(object_name, comment)
    return render_template('party_info.html', row=row)


@bp_party_to_pallet.route('/party_info_issues_to_pallet/<string:row>', methods=['GET', 'POST'])
def party_info_issues_to_pallet(row):
    con = sl.connect('BAZA.db')
    try:
        with con:
            df_all_pallets = pd.read_sql(f"SELECT pallet FROM baza", con).drop_duplicates(subset='pallet', keep='first')
            try:
                df_all_pallets['pallet'] = df_all_pallets['pallet'].fillna(value=np.nan).fillna(0).astype(int)
                df_all_pallets = df_all_pallets.sort_values(by='pallet')
                last_pall_numb = int(df_all_pallets.values[-1].tolist()[0])
                new_pallet_numb = last_pall_numb + 1
            except:
                new_pallet_numb = 0
            print(new_pallet_numb)
            data = pd.read_sql(f"Select parcel_plomb_numb, custom_status_short from baza where party_numb = '{row}'", con)
            data_refuse_plombs = data.loc[data['custom_status_short'] == 'ИЗЪЯТИЕ'].drop_duplicates(subset='parcel_plomb_numb')
            all_plombs = data.drop_duplicates(subset='parcel_plomb_numb').loc[data['parcel_plomb_numb'] != '']
            df_plomb_not_refuse = all_plombs[~all_plombs.parcel_plomb_numb.isin(data_refuse_plombs.parcel_plomb_numb)]
            print(df_plomb_not_refuse)
            for parcel_plomb_numb in df_plomb_not_refuse['parcel_plomb_numb']:
                con.execute(
                    f"Update baza set pallet = '{new_pallet_numb}' where parcel_plomb_numb = '{parcel_plomb_numb}' and VH_status is ?", (None,))
                print(f'{parcel_plomb_numb} - pallet updated')
            df_new_pallet = pd.read_sql(f"Select * from baza where pallet = '{new_pallet_numb}'", con)
            writer = pd.ExcelWriter(f'{addition_folder}Паллет {new_pallet_numb}.xlsx', engine='xlsxwriter')
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
    except Exception as e:
        return {'message': str(e)}, 400

    flash(f'Паллет {new_pallet_numb} c 0-ми из {row} сформирован!', category='success')
    object_name = new_pallet_numb
    comment = f'Паллет c 0-ми из {row} сформирован!'
    insert_user_action(object_name, comment)
    return render_template('party_info.html', row=row)