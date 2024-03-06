import zipfile
import os
import shutil
import glob
import pandas as pd
# Makro do wyciągania nazw punktów z backupów robota wyłącznie z programów procesowych
path='C:\XXX\Backup_FINAL'
excel_path='C:\XXX\points_from_backups.xlsx'
df_points=pd.DataFrame(columns=['Robot','Nazwa punktu','Index'])
df_new_point=pd.DataFrame(columns=['Robot','Nazwa punktu','Index'])
for area_name in os.listdir(path):
    area_path=os.path.join(path,area_name)
    for robot_name in os.listdir(area_path):
        robot_path=os.path.join(area_path,robot_name)
        #print(robot_path)
        file_pattern=f'{robot_path}\\UP7*.ls'
        files=glob.glob(file_pattern)
        #print(files)
        if  not files:
            print('Zjebane')
            print(robot_path)
        for file in files:
            with open(file,'r') as plik:
                content=plik.read()
                pozycja_punktu=content.find('9Y4')
                print( file)
                while pozycja_punktu !=-1:
                    pozycja_programu=0;
                    pozycja_punktu_koniec=content.find(';',pozycja_punktu)
                    if content.find('_0210_',pozycja_punktu,pozycja_punktu_koniec)>0:
                        pozycja_programu = content.find("\"S-Punkt\"=", pozycja_punktu)
                        #print(f'Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]},{content[pozycja_programu + 10:pozycja_programu + 14]}')
                        df_new_point=[robot_name,content[pozycja_punktu:pozycja_punktu_koniec],content[pozycja_programu + 10:pozycja_programu + 14]]
                        df_points.loc[len(df_points)]=df_new_point
                    elif content.find('_1220_', pozycja_punktu, pozycja_punktu_koniec)>0:
                        pozycja_programu = content.find('"ProgNr"=', pozycja_punktu)
                        #print(f'Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]},{content[pozycja_programu+9:pozycja_programu+13]}')
                        df_new_point = [robot_name, content[pozycja_punktu:pozycja_punktu_koniec],content[pozycja_programu + 9:pozycja_programu + 13]]
                        df_points.loc[len(df_points)] = df_new_point
                    elif content.find('_0420_', pozycja_punktu, pozycja_punktu_koniec)>0:
                        pozycja_programu = content.find('"Programm-Nr"=', pozycja_punktu)
                        #print(f'Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]},{content[pozycja_programu+14:pozycja_programu+18]}')
                        df_new_point = [robot_name, content[pozycja_punktu:pozycja_punktu_koniec],content[pozycja_programu + 14:pozycja_programu + 18]]
                        df_points.loc[len(df_points)] = df_new_point
                    elif content.find('_1791_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        pozycja_programu = content.find('"ProgNr"=', pozycja_punktu)
                        #print(f'Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]},{content[pozycja_programu + 9:pozycja_programu + 13]}')
                        df_new_point = [robot_name, content[pozycja_punktu:pozycja_punktu_koniec],content[pozycja_programu + 9:pozycja_programu + 13]]
                        df_points.loc[len(df_points)] = df_new_point
                    elif content.find('_0780_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        pozycja_programu = content.find('"ProgNr"=', pozycja_punktu)
                        #print(f'Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]},{content[pozycja_programu + 9:pozycja_programu + 13]}')
                        df_new_point = [robot_name, content[pozycja_punktu:pozycja_punktu_koniec],content[pozycja_programu + 9:pozycja_programu + 13]]
                        df_points.loc[len(df_points)] = df_new_point
                    elif content.find('_1150_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        pozycja_programu = content.find('"ProgNr"=', pozycja_punktu)
                        #print(f'Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]},{content[pozycja_programu + 9:pozycja_programu + 13]}')
                    elif content.find('_1330_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        pozycja_programu = content.find('"ProgNr"=', pozycja_punktu)
                        #print(f'Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]},{content[pozycja_programu + 9:pozycja_programu + 13]}')
                        df_new_point = [robot_name, content[pozycja_punktu:pozycja_punktu_koniec],content[pozycja_programu + 9:pozycja_programu + 13]]
                        df_points.loc[len(df_points)] = df_new_point
                    elif content.find('_0131_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        pozycja_programu = content.find('"ProgNr"=', pozycja_punktu)
                        #print(f'Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]},{content[pozycja_programu + 9:pozycja_programu + 13]}')
                        df_new_point = [robot_name, content[pozycja_punktu:pozycja_punktu_koniec],0]
                        df_points.loc[len(df_points)] = df_new_point
                    elif content.find('_0135_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        pozycja_programu = content.find('"ProgNr"=', pozycja_punktu)
                        # print(f'Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]},{content[pozycja_programu + 9:pozycja_programu + 13]}')
                        df_new_point = [robot_name, content[pozycja_punktu:pozycja_punktu_koniec], 0]
                        df_points.loc[len(df_points)] = df_new_point
                    else:
                        print(f'!!!!!!!!!!!!!!!Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]}')
                    pozycja_punktu=content.find('9Y4',pozycja_punktu_koniec+1)

#print(df_points)

with pd.ExcelWriter(excel_path) as excel:
    df_points.to_excel(excel, sheet_name='Backup', index=False)
