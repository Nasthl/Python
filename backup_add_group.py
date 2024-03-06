import zipfile
import os
import shutil
import glob
# Makro do dodania grup w backupach robotów

path='C:\\Users\RWSwiss\Downloads\keu1g1111030r01rs1kfu1\ENGLISH\\'
file_pattern=f'{path}UP001.ls'
files=glob.glob(file_pattern)
print(files)

for file in files:
    with open(file, 'r+') as plik:
        zawartosc=plik.read()
        pozycja_gp1=zawartosc.find('GP1:')
        tool_base = zawartosc[pozycja_gp1+6:pozycja_gp1 + 21]
        fragment_gp2 ='GP2:\n'+tool_base+'\n'+'J1=     0.000  mm,	J2=     0.000  mm,	J3=     0.000  mm,'+'\n'+'J4=     0.000  mm,	J5=     0.000  mm,	J6=     0.000  mm,'+'\n'+	'J7=     0.000  mm\n'
        #print(fragment_gp2)
        fragment_gp3 = 'GP3:\n' + tool_base + '\n' + 'J1=     0.000  mm\n'
        #print(fragment_gp3)
        fragment_gp4 = 'GP4:\n' + tool_base + '\n' + 'J1=     0.000  mm\n'
        #print(fragment_gp4)
        fragment_gp5 = 'GP5:\n' + tool_base + '\n' + 'J1=     0.000  mm,	J2=     0.000  mm,	J3=     0.000  mm,' + '\n' + 'J4=     0.000  mm\n'
        #print(fragment_gp5)
        #print(zawartosc[:2000])
        print(tool_base)

        while pozycja_gp1 !=-1:
            pozycja_end=zawartosc.find('};',pozycja_gp1)
            print(f'{file} GP1: {pozycja_gp1} koniec: {pozycja_end}')
            #print(f'Tool i baza:{zawartosc[pozycja_end-1]}')

            pozycja_gp2 = zawartosc.find('GP2:', pozycja_gp1, pozycja_end)
            if pozycja_gp2 == -1:
                print(file, ' Nie ma GP2')
                plik.seek(5013)
                plik.write("KUPA\n"+zawartosc[pozycja_end:])
                #print(zawartosc[pozycja_end:])
               # plik.write(fragment_gp2+fragment_gp3+fragment_gp4+fragment_gp5+zawartosc[pozycja_end-1:])
            #pozycja_gp3 = zawartosc.find('GP3:', pozycja_gp1, pozycja_end)
            # if pozycja_gp3 == -1:
            #     print(file, ' Nie ma GP3')
            # pozycja_gp4 = zawartosc.find('GP4:', pozycja_gp1, pozycja_end)
            # if pozycja_gp4 == -1:
            #     print(file, ' Nie ma GP4')
            # pozycja_gp5 = zawartosc.find('GP5:', pozycja_gp1, pozycja_end)
            # if pozycja_gp5 == -1:
            #     print(file, ' Nie ma GP5')

            pozycja_gp1=zawartosc.find('GP1:',pozycja_gp1+1)


# def dopisz_tekst_do_pliku(nazwa_pliku, fragment_tekstu, pozycja):
#     with open(nazwa_pliku, 'r+') as plik:
#         zawartosc = plik.read()
#         pozycja_znaku = zawartosc.index('\n', pozycja) + 1
#         plik.seek(pozycja_znaku)
#         plik.write(fragment_tekstu + zawartosc[pozycja_znaku:])