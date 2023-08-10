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

df_trigers = pd.read_excel('Triger_dict.xlsx')
print(df_trigers)
triger_dict = df_trigers.set_index('trigger').to_dict()['weight']

triger_list = list(triger_dict.keys())

file_name = filedialog.askopenfilename()
print(file_name)

head, tail = os.path.split(file_name)
print(tail)
df = pd.read_excel(file_name)
name = df[df.columns[19]]
weight = df[df.columns[26]]
parcel_numb = df[df.columns[0]]
good_link = df[df.columns[18]]
good_quont = df[df.columns[21]]
warning_df = pd.DataFrame()
for i in range(0, len(df)):
    print(i)
    good_weight = weight[i] / good_quont[i]
    for trigger in triger_list:
        trigger_weight = triger_dict[trigger]
        if trigger.lower() in name[i].lower():
            if good_weight >= trigger_weight:
                df_to_append = pd.DataFrame({'parcel_numb': [parcel_numb[i]], 'name': [name[i]],
                                             'good_link': [good_link[i]], 'weight': good_weight,
                                             'trigger': [trigger], 'trigger_weight': [trigger_weight], 'class': 0})
                warning_df = pd.concat([warning_df, df_to_append])
        elif weight[i] >= 10:
            df_to_append = pd.DataFrame({'parcel_numb': [parcel_numb[i]], 'name': [name[i]],
                                         'good_link': [good_link[i]], 'weight': good_weight,
                                         'trigger': '', 'trigger_weight': '', 'class': 1})
            warning_df = pd.concat([warning_df, df_to_append])
            break
warning_df = warning_df.sort_values(by=['class', 'weight'])
print(warning_df)
writer = pd.ExcelWriter(f'WARNING_{tail}.xlsx', engine='xlsxwriter')
warning_df.to_excel(writer, sheet_name='Sheet1', index=False)
writer.save()