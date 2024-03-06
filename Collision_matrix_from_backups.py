import zipfile
import os
import shutil
import glob
import pandas as pd
import openpyxl
import os.path
import numpy as np
# Makro do sprawdzenia poprawności sygnałów zon kolizyjnych w backupach i utworzenia collision matrixow
def f_zony_do_coll_matrix(df_coll,robot):
    nowa_df_coll=pd.DataFrame(index=df_coll.index,columns=[''])
    for i in range(len(df_coll)):
        if isinstance(df_coll.iloc[i,0],set):
            temp=df_coll.iloc[i,0]
            for j in temp:
                if j[1]==robot:
                    if not isinstance(nowa_df_coll.iloc[i,0], list):
                        nowa_df_coll.iloc[i, 0] =nowa_df_coll.iloc[i,0]=[nowa_df_coll.iloc[i,0]]
                    nowa_df_coll.iloc[i,0].append(j[0])
            if isinstance(nowa_df_coll.iloc[i, 0], list):
                #print(f'Przed {nowa_df_coll.iloc[i, 0]}')
                nowa_df_coll.iloc[i, 0].pop(0)
                #print(f'PO { nowa_df_coll.iloc[i, 0]}')
    return(nowa_df_coll)

def znajdz_verr_komentarz(file):
    global df_bledy
    with open(file, 'r') as plik:
        content = plik.read()
    pozycja_verr_com = content.find(' Verriegelung ')
    # print( file)
    check_ver = set()
    list_ver = set()
    while pozycja_verr_com != -1:
        pozycja_verr_com_end = content.find(';', pozycja_verr_com + 1)
        # print(content[pozycja_verr_com:pozycja_verr_com_end])
        verr_com = content[pozycja_verr_com + 1:pozycja_verr_com_end]
        verr_com_nr = verr_com[verr_com.find(' ') + 1:verr_com.find(' ', verr_com.find(' ') + 1)]
        verr_com_robot = verr_com[verr_com.find('mit ') + 4:verr_com.find(' ', verr_com.find('mit ') + 5)]
        if not (verr_com_nr, verr_com_robot) in list_ver:
            list_ver.add((verr_com_nr, verr_com_robot))
        if verr_com.find('EIN') > 0:
            if not verr_com_nr in check_ver:
                check_ver.add(verr_com_nr)
            else:
                # print(f"{os.path.basename(file)} komentarz:  podwojne wywolanie EIN zony {verr_com_nr}")
                df_bledy.loc[len(df_bledy), :] = (
                    f"{os.path.basename(file)} komentarz:  podwojne wywolanie EIN zony {verr_com_nr}")
        if verr_com.find('AUS') > 0:
            if verr_com_nr in check_ver:
                check_ver.remove(verr_com_nr)
            else:
                # print(f"{os.path.basename(file)} komentarz: niepotrzebne zwolnienie AUS zony{verr_com_nr}")
                df_bledy.loc[len(df_bledy), :] = (
                    f"{os.path.basename(file)} komentarz: niepotrzebne zwolnienie AUS zony{verr_com_nr}")
        pozycja_verr_com = content.find(' Verriegelung ', pozycja_verr_com_end + 1)
    if not (len(check_ver) == 0):
        for i in check_ver:
            # print(f'{os.path.basename(file)} komentarz:  niezwolniona zona EIN {i}')
            df_bledy.loc[len(df_bledy), :] = (f'{os.path.basename(file)} komentarz:  niezwolniona zona EIN {i}')
    return (list_ver)

def znajdz_wszystkie_DI(content):
    pozycja_sig = content.find('DI[', content.find('P[1]'))
    global df_bledy
    list_sig = []
    while pozycja_sig != -1:
        sig = int(content[pozycja_sig + 3:pozycja_sig + 5])
        check_if_wait = not (content[content.rfind("\n", 0, pozycja_sig) + 11:content.rfind("\n", 0,
                                                                                            pozycja_sig) + 15] == "WAIT")
        if (sig > 40) and (sig < 57) and (not (content[pozycja_sig + 5].isnumeric()) and check_if_wait):
            list_sig.append([sig, pozycja_sig])
        pozycja_sig = content.find('DI[', pozycja_sig + 1)
    return (list_sig)

def usun_poprawne_DI(lista_DI, poczatek, koniec, kolizja):
    global df_bledy
    wynik = []
    for i in lista_DI:
        if not (i[0] == kolizja and i[1] > poczatek and i[1] < koniec):
            wynik.append(i)
    return (wynik)

def znajdz_verr_sygnal(file):
    global df_bledy
    with open(file, 'r') as plik:
        content = plik.read()
    pozycja_verr_sig = content.find('DO[')
    check_ver = set()
    list_ver = set()
    lista_DI = znajdz_wszystkie_DI(content)
    while pozycja_verr_sig != -1:
        sig = int(content[pozycja_verr_sig + 3:pozycja_verr_sig + 5])
        if (sig > 80) and (sig < 97) and (not (content[pozycja_verr_sig + 5].isnumeric())):
            if not sig - 80 in list_ver:
                list_ver.add(sig - 80)
            if not sig in check_ver:
                check_ver.add(sig)
                pozycja_verr_sig_end = content.find(f'DO[{sig - 40}', pozycja_verr_sig + 1)
                while content[pozycja_verr_sig_end + 5].isnumeric():
                    pozycja_verr_sig_end = content.find(f'DO[{sig - 40}', pozycja_verr_sig_end + 1)
                    print(f'{file} Wszedlem do petli')
                if pozycja_verr_sig_end != -1:
                    sprawdz_DI_kolizji(content, sig, pozycja_verr_sig, pozycja_verr_sig_end)
                    lista_DI = usun_poprawne_DI(lista_DI, pozycja_verr_sig, pozycja_verr_sig_end, sig - 40)
            else:
                # print(f'{os.path.basename(file)} Podwojne wywolanie zony {sig-80}')
                df_bledy.loc[len(df_bledy), :] = (f'{os.path.basename(file)} Podwojne wywolanie zony {sig - 80}')
        if (sig > 40) and (sig < 57) and (not (content[pozycja_verr_sig + 5].isnumeric())):
            if not sig - 40 in list_ver:
                list_ver.add(sig - 40)
            if sig + 40 in check_ver:
                check_ver.remove(sig + 40)
            else:
                # print(f'{os.path.basename(file)} Brak sygnalu wywolania zony {sig-40}')
                df_bledy.loc[len(df_bledy), :] = (f'{os.path.basename(file)} Brak sygnalu wywolania zony {sig - 40}')
        pozycja_verr_sig = content.find('DO[', pozycja_verr_sig + 1)

    if len(check_ver) > 0:
        for i in check_ver:
            # print(f'{os.path.basename(file)} Brak sygnalu zwolnienia zona {i - 80}')
            df_bledy.loc[len(df_bledy), :] = (f'{os.path.basename(file)} Brak sygnalu zwolnienia zona {i - 80}')
    if len(lista_DI) > 0:
        for i in lista_DI:
            poz_pop_punktu = content.rfind('P[', 0, i[1])
            # print(f'{os.path.basename(file)} Niepotrzebny sygnal: {i[0]} w punkcie {content[poz_pop_punktu:poz_pop_punktu+5]} ')
            df_bledy.loc[len(df_bledy), :] = (
                f'{os.path.basename(file)} Niepotrzebny sygnal: {i[0]} w punkcie {content[poz_pop_punktu:poz_pop_punktu + 5]} ')
    return (list_ver)

def sprawdz_DI_kolizji(zawartosc, kolizja, pozycja_poczatek, pozycja_koniec):
    global df_bledy
    pozycja_p=0
    while pozycja_poczatek < pozycja_koniec and pozycja_poczatek != -1 and pozycja_p!=-1:
        pozycja_p = zawartosc.find(' P[', pozycja_poczatek)
        pozycja_tc = zawartosc.find('TC_ONLINE', pozycja_poczatek)
        check_sig = zawartosc[zawartosc.find(f'DI[{kolizja - 40}', pozycja_tc, pozycja_p) + 5]
        pozycja_poczatek = zawartosc.find('TC_ONLINE', pozycja_p + 1)
        if zawartosc[pozycja_poczatek + 11] == 'O':
            pozycja_poczatek = zawartosc.find('TC_ONLINE', pozycja_poczatek + 1)


        if not ((zawartosc.find(f'DI[{kolizja - 40}', pozycja_tc, pozycja_p) != -1) or (check_sig.isnumeric())):
            poz_pop_punktu = zawartosc.rfind(' P[', 0, pozycja_p)
            #print('a tu?')
           # print(zawartosc.find(f'DI[{kolizja - 40}', pozycja_tc, pozycja_p))
            #print(check_sig.isnumeric())
           # print('przed', pozycja_p, ' ', pozycja_poczatek, ' ', pozycja_koniec, ' ', zawartosc[pozycja_tc:pozycja_p])
            #print(f'{os.path.basename(file)} W TC_ONLINE w punkcie {zawartosc[poz_pop_punktu:poz_pop_punktu + 6]}, brak sygnalu {kolizja - 40}')
            df_bledy.loc[len(df_bledy), :] = (f'{os.path.basename(file)} W TC_ONLINE w punkcie {zawartosc[poz_pop_punktu:poz_pop_punktu + 6]}, brak sygnalu {kolizja - 40}')

    return (0)


#path = 'C:\Porsche PO546\Backup_test'
#path = 'C:\Porsche PO546\Backup_FINAL'
path = 'C:\\Users\\tomasz.korzyniec\\Downloads\\backups'
for area_name in os.listdir(path):
    excel_path = f'C:\Porsche PO546\\{area_name}.xlsx'
    area_path = os.path.join(path, area_name)
    d_verr_matrix = {}
    if not (os.path.isfile(excel_path)):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = 'Podsumowanie'
        wb.save(excel_path)
    with pd.ExcelWriter(excel_path) as excel:
        df_matrix = pd.DataFrame(columns=os.listdir(area_path), index=range(len(os.listdir(area_path))))
        df_matrix.index = os.listdir(area_path)
        d_robot_ver_matrix = {}
        # df_matrix.insert(loc=0, column='Robot\Robot', value=os.listdir(area_path))
        for robot_name in os.listdir(area_path):
            df_bledy = pd.DataFrame(columns=['Błędy'])
            df_zony = pd.DataFrame(columns=['Zony'])
            df_zony_all = pd.DataFrame(columns=['Wszystkie zony'])
            df_matrix.loc[robot_name, robot_name] = 'X'
            robot_path = os.path.join(area_path, robot_name)
            file_pattern = f'{robot_path}\\UP*.ls'
            files = glob.glob(file_pattern)
            files=[file for file in files if file!=f'{robot_path}\\updtlog.ls']
            print(files)
            names = ['']
            for file in files:
                name = os.path.basename(file)
                names.append(name)
            df_ver_matrix = pd.DataFrame(columns=[''], index=range(len(files) + 1))
            df_ver_matrix.index = names
            lista_ver_robot = set()
            if not files:
                print('Zjebane')
                print(robot_path)
            else:
                for file in files:
                    print(f'{area_name} {robot_name} {file}')
                    lista_ver_com = znajdz_verr_komentarz(file)
                    lista_ver_com_sig = {}
                    if len(lista_ver_com) > 0:
                        # print(f'{os.path.basename(file)} {lista_ver_com}')
                        df_zony.loc[len(df_zony), :] = (f'{os.path.basename(file)} {lista_ver_com}')
                        # df_ver_matrix.loc[len(df_zony),0]=os.path.basename(file)
                        df_ver_matrix.loc[os.path.basename(file), :] = lista_ver_com
                        lista_ver_com_sig = {int(tupla[0]) for tupla in lista_ver_com}
                    else:
                        df_zony.loc[len(df_zony), :] = (f'{os.path.basename(file)}  brak zon kolizyjnych')
                    lista_ver_sig = znajdz_verr_sygnal(file)
                    if len(set(lista_ver_sig) - set(lista_ver_com_sig)) > 0:
                        # print(f'{os.path.basename(file)} Wiecej zon w sygnalach niz w komentarzach {set(lista_ver_sig)-set(lista_ver_com_sig)}')
                        df_bledy.loc[len(df_bledy), :] = (
                            f'{os.path.basename(file)} Wiecej zon w sygnalach niz w komentarzach {set(lista_ver_sig) - set(lista_ver_com_sig)}')
                    if len(set(lista_ver_com_sig) - set(lista_ver_sig)) > 0:
                        # print(f'{os.path.basename(file)} Wiecej zon w komentarzach niz w sygnalach {set(lista_ver_com_sig)-set(lista_ver_sig)}')
                        df_bledy.loc[len(df_bledy), :] = (
                            f'{os.path.basename(file)} Wiecej zon w komentarzach niz w sygnalach {set(lista_ver_com_sig) - set(lista_ver_sig)}')
                    lista_ver_robot.update(lista_ver_com)
                # slownik do przypisania do robota zon kolizyjnych dla kazdej UPy
                d_robot_ver_matrix[robot_name] = df_ver_matrix
                df_zony_all.loc[len(df_zony_all), :] = (f'Lista kolizji dla {robot_name} to {lista_ver_robot}')
                for i in lista_ver_robot:
                    if i[1] in df_matrix.columns and not (pd.isna(df_matrix.loc[robot_name, i[1]])):
                        kolizje = i[0] + ' ' + df_matrix.loc[robot_name, i[1]]
                    else:
                        kolizje = i[0]
                    df_matrix.loc[robot_name, i[1]] = kolizje
                    for j in lista_ver_robot:
                        if i != j:
                            if i[0] == j[0]:
                                # print(f'{file} Podwójna zona: {i} oraz {j}')
                                df_zony_all.loc[len(df_zony_all), :] = (
                                    f'{os.path.basename(file)} Ten sam nr zony dla dwóch różnych robotów {i} oraz {j}')
                df_zony_all.to_excel(excel, sheet_name=robot_name, index=False)
                df_zony.to_excel(excel, sheet_name=robot_name, index=False, startrow=len(df_zony_all) + 1)
                df_bledy.to_excel(excel, sheet_name=robot_name, index=False,
                                  startrow=len(df_zony_all) + 1 + len(df_zony) + 1)
        df_matrix.to_excel(excel, sheet_name='podsumowanie')
        #excel.book.move_sheet('podsumowanie', offset=-len(os.listdir(area_path)))
    # print(d_robot_ver_matrix)
    for d_robot, d_verr in d_robot_ver_matrix.items():
        excel_path = f'C:\Porsche PO546\\Verrigelung_matrix\\{area_name}_{d_robot}.xlsx'
        with pd.ExcelWriter(excel_path) as excel:
            for d_robot2, d_verr2 in d_robot_ver_matrix.items():
                if not (d_robot == d_robot2):
                    d_verr_temp =f_zony_do_coll_matrix(d_verr,d_robot2)
                    d_verr2_temp = f_zony_do_coll_matrix(d_verr2, d_robot)
                    d_verr_t = d_verr2_temp.T
                    temp=np.empty((len(d_verr_temp),len(d_verr2_temp)), dtype=object)
                    df_r_vs_r_matrix=pd.DataFrame(temp)
                    for i in range(len(d_verr_temp)):
                        for j in range(len(d_verr2_temp)):
                            if isinstance(d_verr_temp.iloc[i,0], list) and isinstance(d_verr2_temp.iloc[j,0], list):
                                #print(f'{type(d_verr_temp.iloc[i,0])} {type(d_verr2_temp.iloc[j,0])}')
                                wspolne=list(set(d_verr_temp.iloc[i,0])&set(d_verr2_temp.iloc[j,0]))
                                if len(wspolne)>0:
                                    df_r_vs_r_matrix.iloc[i,j]=wspolne
                    temp=np.empty((2,2),dtype=object)
                    temp[0,1]=d_robot2
                    temp[1,0]=d_robot
                    temp[1,1]='Collisions'
                    df_naglowek=pd.DataFrame(temp)
                    df_r_vs_r_matrix.to_excel(excel, sheet_name=f'{d_robot}_{d_robot2}', index=False, startcol=1)
                    styler=d_verr_temp.style.set_properties(**{'background-color': 'yellow'})
                    styler.to_excel(excel, sheet_name=f'{d_robot}_{d_robot2}')
                    styler = d_verr_t.style.set_properties(**{'background-color': 'yellow'})
                    styler.to_excel(excel, sheet_name=f'{d_robot}_{d_robot2}')
                    df_naglowek.to_excel(excel, sheet_name=f'{d_robot}_{d_robot2}',index=False, header=False)