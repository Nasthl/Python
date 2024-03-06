import os
import glob
import datetime
import shutil
# Makro do wyciągnięcia z plików xvr informacji nt parametrów safe zone  

path='C:\\Users\\tomasz.korzyniec\\Desktop\\09'
#excel_path='C:\\Users\\tomasz.korzyniec\\Desktop\\09\\summary.xlsx'


for paket_name in os.listdir(path):
    paket_path=os.path.join(path,paket_name)
    for area_name in os.listdir(paket_path):
        area_path=os.path.join(paket_path,area_name)
        file_pattern = f'{area_path}\\*.xvr'
        files = glob.glob(file_pattern)
        sum_coordinates = 0
        for file in files:

            with open(file, 'r') as plik:
                content = plik.read()
                flaga=True
                pozycja_group_user_frame = content.find('<VAR name="$DCSS_UFRM">')
                if pozycja_group_user_frame!=-1:
                    pozycja_group_user_frame_koniec = content.find('</VAR>', pozycja_group_user_frame)
                    pozycja_user_frame=content.find('<ARRAY name="$DCSS_UFRM[1,')
                    while pozycja_user_frame != -1:
                        pozycja_user_frame_koniec = content.find('</ARRAY>',pozycja_user_frame)
                        pozycja_wspolrzednych=content.find('value="',content.find('<FIELD name="$X"',pozycja_user_frame))
                        sum_coordinates=0
                        while pozycja_wspolrzednych<pozycja_user_frame_koniec:
                            sum_coordinates+=abs(float(content[pozycja_wspolrzednych+7:pozycja_wspolrzednych+12]))
                            #if sum_coordinates>0:
                                #print(f'{file} user frame WIEKSZE OD ZERA')
                               # flaga = False
                               # break
                            pozycja_wspolrzednych=content.find('value="',pozycja_wspolrzednych+1)
                        #print('Nowy')
                        pozycja_user_frame = content.find('<ARRAY name="$DCSS_UFRM[1,',pozycja_user_frame+1)
                        #if sum_coordinates > 0:
                          #  flaga = False
                            #break
                pozycja_group_safe_zone = content.find('<VAR name="$DCSS_CPC">')
                liczba_zon=0
                iterator = 0
                wspolrzedne=[]
                if pozycja_group_safe_zone != -1:
                    pozycja_group_safe_zone_koniec = content.find('</VAR>', pozycja_group_safe_zone)
                    pozycja_safe_zone = content.find('<ARRAY name="$DCSS_CPC[')
                    while pozycja_safe_zone != -1:
                        pozycja_safe_zone_koniec = content.find('</ARRAY>', pozycja_safe_zone)
                        pozycja_wspolrzednych = content.find(']" value="',content.find('<FIELD name="$X"', pozycja_safe_zone))
                        sum_coordinates = 0

                        while pozycja_wspolrzednych < pozycja_safe_zone_koniec and pozycja_wspolrzednych!=-1 :
                            #print(content[pozycja_wspolrzednych + 10:pozycja_wspolrzednych + 15])
                            sum_coordinates += abs(float(content[pozycja_wspolrzednych + 10:pozycja_wspolrzednych + 15]))

                            if liczba_zon==0:
                                wspolrzedne.append(float(content[pozycja_wspolrzednych + 10:pozycja_wspolrzednych + 15]))
                                sum_coordinates=0
                                iterator += 1
                            if sum_coordinates > 0 and liczba_zon>0:
                                #print(f'{file}  safe zone WIEKSZE OD ZERA')
                                flaga = False
                                break
                            pozycja_wspolrzednych = content.find(']" value="', pozycja_wspolrzednych + 1)
                        liczba_zon+=1
                        #print(f'{sum_coordinates} {liczba_zon} {content[pozycja_safe_zone:pozycja_safe_zone_koniec]}')
                        pozycja_safe_zone = content.find('<ARRAY name="$DCSS_CPC[', pozycja_safe_zone_koniec + 1)
                        if sum_coordinates > 0:
                            flaga=False
                            break
                if not(flaga):
                #if flaga and iterator<=8:
                    #if len(set(wspolrzedne))==4:
                       # DI_check=True
                    #else:
                        #DI_check=False
                    print(file, ' ', iterator,'   ', wspolrzedne)
                flaga=True
                #print(content[pozycja_group_safe_zone:pozycja_group_safe_zone_koniec])

