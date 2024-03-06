import os
import shutil
import glob
import pandas as pd
# Makro to wypisania nazw punktów procesowych z backupów wraz z ich współrzędnymi dla różnych technologii

def znajdz_koordynaty(zawartosc,start):
    nazwa_punktu=zawartosc[zawartosc.find(' P[',start)+1:zawartosc.find(']',zawartosc.find(' P[',start))+1]
    loc_wspolrzedne=zawartosc.find(nazwa_punktu+'{')
    loc_x=zawartosc.find('X =',loc_wspolrzedne)
    kor_x=zawartosc[loc_x+4:zawartosc.find('mm',loc_x)]
    loc_y = zawartosc.find('Y =', loc_wspolrzedne)
    kor_y = zawartosc[loc_y+4:zawartosc.find('mm', loc_y)]
    loc_z = zawartosc.find('Z =', loc_wspolrzedne)
    kor_z = zawartosc[loc_z+4:zawartosc.find('mm', loc_z)]

    #print(nazwa_punktu)
    #print(zawartosc[loc_wspolrzedne:loc_wspolrzedne+100])
    #print(f'{kor_x} {kor_y} {kor_z}')
    return kor_x,kor_y,kor_z

path='C:\\Users\\tomasz.korzyniec\\Downloads\\backups'
excel_path='C:\Porsche PO546\points_from_backups.xlsx'
df_points=pd.DataFrame(columns=['Area','Robot','UP','Nazwa punktu','Index','X','Y','Z'])
df_new_point=pd.DataFrame(columns=['Area','Robot','UP','Nazwa punktu','Index','X','Y','Z'])
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
                up_name = os.path.basename(file)
                print( file)
                while pozycja_punktu !=-1:
                    pozycja_programu=0;
                    pozycja_punktu_koniec=content.find(' ;',pozycja_punktu)
                    #zmienna do sprawdzenia czy znaleziono punkt procesowy
                    check_if_point_exist=0
                    if content.find('_0210_',pozycja_punktu,pozycja_punktu_koniec)>0:
                        index_name='"ProgNo"='
                        check_if_point_exist = 1
                    elif content.find('_1220_', pozycja_punktu, pozycja_punktu_koniec)>0:
                        index_name = '"ProgNo"='
                        check_if_point_exist = 1
                    elif content.find('_0420_', pozycja_punktu, pozycja_punktu_koniec)>0:
                        index_name = '"ProgNo"='
                        check_if_point_exist = 1
                    elif content.find('_1791_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        index_name = '"ProgNo"='
                        check_if_point_exist = 1
                    elif content.find('_0780_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        index_name = '"ProgNo"='
                        check_if_point_exist = 1
                    elif content.find('_1150_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        index_name = '"ProgNo"='
                        #check_if_point_exist = 1
                    elif content.find('_1330_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        index_name = '"ProgNo"='
                    elif content.find('_0131_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        index_name = '"ProgNo"='
                        check_if_point_exist = 1
                    elif content.find('_0135_', pozycja_punktu, pozycja_punktu_koniec) > 0:
                        index_name = '"ProgNo"='
                        check_if_point_exist = 1
                    else:
                        print(f'!!!!!!!!!!!!!!!Punkt to {content[pozycja_punktu:pozycja_punktu_koniec]}')
                    #Jesli znalazlo punkt to znajdz jego koordynaty i wpisz do df
                    if check_if_point_exist:
                        pozycja_programu = content.find(index_name, pozycja_punktu)
                        x, y, z = znajdz_koordynaty(content, pozycja_punktu)
                        # Sprawdz czy na pewno jest poprawna nazwa do wyszukania indexu - zobacz czy roznica miedzy pozycja_programu, a przecinkiem nie jest jakas za duza
                        if ((content.find(',', pozycja_programu) - pozycja_programu) < 20) and pozycja_programu > 0:
                            df_new_point = [area_name, robot_name, up_name,
                                            content[pozycja_punktu:pozycja_punktu_koniec],
                                            content[pozycja_programu + len(index_name):content.find(',', pozycja_programu)], x, y, z]
                        else:
                            df_new_point = [area_name, robot_name, up_name,
                                            content[pozycja_punktu:pozycja_punktu_koniec], 'Błąd', x, y, z]
                        print(df_new_point)
                        df_points.loc[len(df_points)] = df_new_point
                    pozycja_punktu = content.find('9Y4', pozycja_punktu_koniec + 1)

#print(df_points)

with pd.ExcelWriter(excel_path) as excel:
    df_points.to_excel(excel, sheet_name='Backup', index=False)
