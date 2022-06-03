# -*- coding: cp1251 -*-
import fdb
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import ttk
from tkinter import *

way2 = 'd:/Basa/STAT2P.fb'

con = fdb.connect(dsn='127.0.0.1:' + way2, user='sysdba', password='masterkey')
cur = con.cursor()

cur.execute("""select product 
            from stat2p 
            group by product""")
por = cur.fetchall()
print(por)
con.commit()

param = ['GrQuantity1', 'GrQuantity2', 'Upormean', 'Rcumean', 'GrNumber17', 'GrNumber18', 'GrNumber19', 'GrNumber21', 'GrNumber23', 'GrNumber25']

def plot2p():
    product1 = comboExample.get()
    param = param1.get()
    way1 = way.get()

    if way1 == '':
        way1 = 'd:/Basa/STAT2P.fb'

    data1 = name_entry.get()
    data2 = surname_entry.get()
    s = len(data1)
    s1 = len(data2)
    if s == 0 and s1 == 0:
        poisk = ''
    elif (s > 0) and (s1 == 0):
        poisk = """and datetimebeg > '""" + data1 +"""'"""
    elif (s == 0) and (s1 > 0):
        poisk = """and datetimebeg < '""" + data2 +"""'"""
    elif (s > 0) and (s1 > 0):
        poisk = """and datetimebeg > '""" + data1 +"""' and datetimebeg < '""" + data2 +"""'"""

    party1 = party.get()
    s2 = len(party1)
    if s2 == 0:
        party2 = ''
    elif (s2 > 0):
        party2 = """and PARCELNUMBER = '""" + party1 + """'"""

    #con = fdb.create_database("create database '127.0.0.1:d:/basa/stat2P.fb' user 'sysdba' password 'masterkey'")       #создаем базу
    #con = fdb.connect(host='127.0.0.1', database=r'd:/Basa/STAT2P.fb', user='sysdba', password='masterkey', charset='UTF8', fb_library_name=r'c:\Program Files\Firebird\Firebird_2_5\WOW64\fbclient.dll')
    con = fdb.connect(dsn='127.0.0.1:' + way1, user='sysdba', password='masterkey')
    cur = con.cursor()

    cur.execute("""select """ + param + """ 
            from stat2p 
            where product = '""" + product1 + """' """ + poisk + """ """ + party2 + """
             order by DATETIMEBEG""")
    por = cur.fetchall()

    con.commit()

    fig, ax = plt.subplots(figsize =(10, 5))
    index = list(por)
    quantity = list(range(len(por)))
    plt.plot(quantity, index, marker='o')    #построение гистограммы
    plt.xlabel('Кол-во пл-н', fontsize=24)
    plt.ylabel(param, fontsize=24)

    ax.grid()
    plt.show()

window = tk.Tk()
window.title("Добро пожаловать в приложение PythonRu")
window.geometry('400x250')
lbl = tk.Label(window, text="Список изделий 2П")
lbl.grid(row=0,column=0)
comboExample = ttk.Combobox(window, values=por)

comboExample.current(0)
comboExample.grid(row=0,column=1)

lbl1 = tk.Label(window, text="""Список параметров""")
lbl1.grid(row=2,column=0)

param1 = ttk.Combobox(window, values=param)
param1.current(0)
param1.grid(row=2,column=1)

lbl2 = tk.Label(window, text="""Поиск по дате""")
lbl2.grid(row=4,column=1)

name = StringVar()
surname = StringVar()

name_label = Label(text="C:")
surname_label = Label(text="До:")

name_label.grid(row=5,column=0)
surname_label.grid(row=6,column=0)

name_entry = Entry(textvariable=name)
surname_entry = Entry(textvariable=surname)
name_entry.grid(row=5,column=1)
surname_entry.grid(row=6,column=1)

lbl3 = tk.Label(window, text="""Поиск по №партии""")
lbl3.grid(row=7,column=0)
party = tk.Entry(window, width=50)
party.grid(row=7,column=1)

lbl = tk.Label(window, text="Путь к базе")
lbl.grid(row=9,column=0)


way = tk.Entry(window, width=50)
way.grid(row=9,column=1)

btn = tk.Button(window, text="Запуск!", command=plot2p)

btn.grid(row=10,column=1)

window.mainloop()