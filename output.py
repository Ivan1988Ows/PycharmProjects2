# -*- coding: cp1251 -*-
import os
from openpyxl import load_workbook
import fdb

n = input()
path = 'd:/Работа/file/2021'
text_files = [f for f in os.listdir(path) if f.endswith('.sta') and f.startswith(n)]  # поиск текстовых файлов в папке path
print(text_files)


#con = fdb.create_database("create database '127.0.0.1:d:/basa/stat2P.fb' user 'sysdba' password 'masterkey'")       #создаем базу
#con = fdb.connect(host='127.0.0.1', database=r'd:/basa/STAT2P.fb', user='sysdba', password='masterkey', charset='UTF8', fb_library_name=r'c:\Program Files\Firebird\Firebird_2_5\WOW64\fbclient.dll')
con = fdb.connect(dsn='127.0.0.1:d:/basa/STAT2P.fb', user='sysdba', password='masterkey')
cur = con.cursor()
cur.execute('''recreate table stat20_21 (Product varchar(30), DateTimeBeg date, ParcelNumber varchar(30), WaferNumber integer, GrQuantity1 integer,
    GrQuantity2 integer, Upormin numeric(5,2), Upormean numeric(5,2), Upormax numeric(5,2), Rcumin numeric(5,2),
     Rcumean numeric(5,2), Rcumax numeric(5,2))''')         #добавляем таблицу
con.commit()




def report(way):
    with open('d:/Работа/file/2021/' + way) as file:    #открытие файла с данными
        data = {}  # создание пустого списка
        data.setdefault('Product')
        data.setdefault('DateTimeBeg')
        data.setdefault('ParcelNumber')
        data.setdefault('WaferNumber')
        data.setdefault('GrQuantity1')
        data.setdefault('GrQuantity2')
        data.setdefault('Upormin')
        data.setdefault('Upormean')
        data.setdefault('Upormax')
        data.setdefault('Rcumin')
        data.setdefault('Rcumean')
        data.setdefault('Rcumax')

        R1 = '<TestName>=Rси 1'
        R2 = '<TestName>=Rси1'
        R3 = '<TestName>=Rси;<Mode>=1,00'
        R4 = '<TestName>=Rси;<Mode>=2,00'
        R5 = '<TestName>=Rси;<Mode>=3,00'
        R6 = '<TestName>=Rси       ;'

        for line in file:
            if '<Product>=' in line:        # Выводим наименование изделия
                test = line.replace(';', '').split('=')
                a = test[1].rstrip()
                data['Product'] = a
            elif '<DateTimeBeg>=' in line:      #дата замера
                test = line.replace(';', '').split('=')
                b = test[1].split()
                d = b[0]
                data['DateTimeBeg'] = d
            elif '<ParcelNumber>=' in line:     #№партии
                test = line.replace(';', '').split('=')
                b1 = test[1].rstrip()
                data['ParcelNumber'] = b1
            elif '<WaferNumber>=' in line:       #№пластины
                test = line.replace(';', '').split('=')
                a1 = test[1].rstrip()
                data['WaferNumber'] = a1

            elif '<TestName>=Uпор' in line:         #порог мин сред макс
                test = line.split(';')
                a3 = test[6].split('=')[1]
                a4 = test[8].split('=')[1]
                a5 = test[7].split('=')[1]
                data['Upormin'] = a3.replace(',', '.')
                data['Upormean'] = a4.replace(',', '.')
                data['Upormax'] = a5.replace(',', '.')



            elif (R1 in line) or (R2 in line) or (R3 in line) or (R4 in line) or (R5 in line) or (R6 in line):
                test = line.split(';')
                a6 = test[6].split('=')[1]
                a7 = test[8].split('=')[1]
                a8 = test[7].split('=')[1]
                data['Rcumin'] = a6.replace(',', '.')
                data['Rcumean'] = a7.replace(',', '.')
                data['Rcumax'] = a8.replace(',', '.')

            try:
                if '<GrNumber>=2;<GrQuantity>=' in line:  # кол-во годных
                    test = line.replace(';', '').split('=')
                    a9 = test[2].rstrip()
                    data['GrQuantity2'] = a9
            except:
                data['GrQuantity2'] = 0
            try:
                if '<GrNumber>=1;<GrQuantity>=' in line:  # кол-во годных
                    test = line.replace(';', '').split('=')
                    a2 = test[2].rstrip()
                    data['GrQuantity1'] = a2
            except:
                data['GrQuantity1'] = 0


        print(data)

        data1 = list(data.values())
        print(data1)


        #wb = load_workbook('stat2p790.xlsx')    #открываем существующий файл
        #ws = wb.active
        #ws.append(data)                 #добавляем ряд данных

        #wb.save('stat2p790.xlsx')
        #wb.close()
        cur = con.cursor()
        cur.execute('''insert into stat20_21 (Product, DateTimeBeg, ParcelNumber, WaferNumber, GrQuantity1, GrQuantity2,
                     Upormin, Upormean, Upormax, Rcumin, Rcumean, Rcumax) values (?,?,?,?,?,?,?,?,?,?,?,?)''', data1)




for name in text_files:
    report(name) #запуск функции report

con.commit()