import pandas as pd
import numpy as np
import openpyxl as xl
import os
import natsort
import glob as gb
from openpyxl import Workbook
import filecmp
import time
 
 
os.getcwd()
 
file_paths_1 = {}    #RS2.8
 
file_paths_2 = {}     #RS2.7
 
check_mass = []
 
#sort_nass = natsort.natsorted(nass_1, reverse=False)
for dirpath, dirnames, filenames in os.walk("."):
    for filename in filenames:
        if filename.endswith('2.8.xlsm'):
          if filename not in file_paths_1:
            file_paths_1[filename] = []
          file_paths_1[filename].append('./' + dirpath.partition('\\')[2] +'/')
 
 
for dirpath, dirnames, filenames in os.walk("."):
    for filename in filenames:
        if filename.endswith('2.7.xlsm'):
          if filename not in file_paths_2:
            file_paths_2[filename] = []
          file_paths_2[filename].append('./' + dirpath.partition('\\')[2] +'/')
 
for file1, path1 in file_paths_1.items():
        check_mass.append(file1.partition('_')[0])
print(set(check_mass))
 
print(len(set(check_mass)))
 
count = 0
for fil2, path2 in file_paths_2.items():
    if fil2.partition('_')[0] in check_mass:
        print("YES")
    else:
        print("NO", fil2)
        count += 1
print(count)
 
 
for file, path in file_paths_1.items():
            df_main = pd.read_excel(path[0] + file, sheet_name='Лист1')
            df_main = df_main.drop(columns=df_main.iloc[:, 1:43])
            k = 1
            for p in path:
                df_temp = pd.read_excel(p + file, sheet_name='Лист1')
                df_temp = df_temp.iloc[:, 1]
                df_main = pd.concat([df_main, df_temp], axis=1)
                # print(k, p, file)
 
                k += 1
            for file2, path2 in file_paths_2.items():
                # print(path[0] + file)
                # print(len(path))
                k = 1
                # print(file, file2, p2, i)
                # i += 1
                if file.partition('_')[0] == file2.partition('_')[0]:  # TODO if file != file2 ===>>>>>    upgrade__
                    # ___new df_main(path2[0] + file2, .....)___
                    # ___Save files (f'output{file}.xlsx')
                    for p2 in path2:
                        df_temp = pd.read_excel(p2 + file2, sheet_name='Лист1')
                        if p2 != './05.2020-10.2021 2.7 new/':
                            df_temp = df_temp.iloc[:, 1: 38]
                        else:
                            df_temp = df_temp.iloc[:, 1: 16]
                        print(df_temp)
                        df_main = pd.concat([df_main, df_temp], axis=1)
                        print(k, p, file, p2, file2)
                        print(file.partition('_')[0], file2.partition('_')[0], p2)
                        k += 1
                        #print(file.partition('_')[0], file2.partition('_')[0], p)
            #if file.partition('_')[0] == file2.partition('_')[0]:  # TODO if file != file2 ===>>>>>    upgrade__
            # ___new df_main(path2[0] + file2, .....)___
            # ___Save files (f'output{file}.xlsx')
 
            toc = time.perf_counter()
            print('Time is: ', toc, f'for {file} ')
            writer = pd.ExcelWriter(f"{file.partition('_')[0]}.xlsx", engine='xlsxwriter',
                            datetime_format="DD-MM-YYYY")
            df_main.to_excel(writer, sheet_name='Лист1', index=False, header=True)
            # df_main.to_excel(f'output{file.partition(".xlsm")[0]}.xlsx', index=False, header=False, formatters={'cost':'{:,.2f}'.format})
            workbook = writer.book
            worksheet = writer.sheets['Лист1']
            format1 = workbook.add_format({'num_format': '#,##0.00'})
            worksheet.set_column(1, 168, 18, format1)
            writer.save()
 
 
for file2, path2 in file_paths_2.items():
    if file2 not in check_mass:
        df_main = pd.read_excel(path2[0] + file2, sheet_name='Лист1')
        df_main = df_main.drop(columns=df_main.iloc[:, 1:43])
        k = 1
        for p2 in path2:
            df_temp = pd.read_excel(p2 + file2, sheet_name='Лист1')
            if p2 != './05.2020-10.2021 2.7 new/':
                df_temp = df_temp.iloc[:, 1: 38]
            else:
                df_temp = df_temp.iloc[:, 1: 16]
            df_main = pd.concat([df_main, df_temp], axis=1)
            print(k, p2, file2)
            k += 1
            # print(file.partition('_')[0], file2.partition('_')[0], p)
    # if file.partition('_')[0] == file2.partition('_')[0]:  # TODO if file != file2 ===>>>>>    upgrade__
    # ___new df_main(path2[0] + file2, .....)___
    # ___Save files (f'output{file}.xlsx')
 
    toc = time.perf_counter()
    print('Time is: ', toc, f'for {file2} ')
    writer = pd.ExcelWriter(f"{file2.partition('_')[0]}.xlsx", engine='xlsxwriter',
                            datetime_format="DD-MM-YYYY")
    df_main.to_excel(writer, sheet_name='Лист1', index=False, header=True)
    workbook = writer.book
    worksheet = writer.sheets['Лист1']
    format1 = workbook.add_format({'num_format': '#,##0.00'})
    worksheet.set_column(1, 168, 18, format1)
    writer.save()
