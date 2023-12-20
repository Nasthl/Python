import random

def stworz_plansze():
    "Tworzy plansze do gry w kółko i krzyżyk."
    plansza = []
    for i in range(7):
        if i % 2 == 0:
            plansza.append(['-','-', '-', '-', '-', '-','-'])
        else:
            plansza.append(['|', ' ', '|', ' ', '|', ' ','|'])
    return plansza

def wyswietl_plansze(plansza):
    "Wyświetla plansze do gry w kółko i krzyżyk."
    for i in range(7):
        for j in range(7):
            print(plansza[i][j], end='')
        print()

def sprawdz_czy_wygrana(plansza):
    "Sprawdza czy jest wygrana."
    wygrany=-1
    for i in range (3):
        k=2*i+1
        if plansza[k][1] == plansza[k][3] == plansza[k][5] and (plansza[k][1]=='X' or plansza[k][1]=='O') :
            wygrany=plansza[k][1]
        if plansza[1][k] == plansza[3][k] == plansza[5][k] and (plansza[1][k] == 'X' or plansza[k][1]=='O'):
            wygrany=plansza[1][k]
    if plansza[1][1] == plansza[3][3] == plansza[5][5] and (plansza[1][1] == 'X' or plansza[1][1]=='O'):
            wygrany=plansza[1][1]
    if plansza[1][5] == plansza[3][3] == plansza[5][1] and (plansza[1][5] == 'X' or plansza[1][5]=='O'):
            wygrany = plansza[1][1]
    if wygrany=='X' or wygrany=='O':
        print('Wygrał gracz', wygrany)
        return 1
    elif not(any(' ' in wiersz for wiersz in plansza)):
        print('Remis')
        return 1
    else:
        return 0

def wpisz_znak(plansza, X, Y, znak):
    "Wpisuje znak na plansze."
    if X>3 or X<1 or Y>3 or Y<1:
        raise ValueError
    if plansza[2*X-1][2*Y-1]==' ':
        plansza[2*X-1][2*Y-1]=znak
    else:
        print('To pole jest już zajęte, tura przepadła')
    return plansza

def tura_komputera(plansza):
    "Ruch komputera."
    puste_pola=[]
    for i in range(1,4):
        for j in range(1,4):
            if plansza[2*i-1][2*j-1]==' ':
               puste_pola.append([i,j])
    losuj_pole=random.randint(0,len(puste_pola)-1)
    plansza=wpisz_znak(plansza,puste_pola[losuj_pole][0],puste_pola[losuj_pole][1],'O')
    return plansza

plansza=stworz_plansze()
wyswietl_plansze(plansza)
while not(sprawdz_czy_wygrana(plansza)):
    X=input('Podaj numer wiersza: ')
    Y=input('Podaj numer kolumny: ')
    try:
        plansza=wpisz_znak(plansza,int(X),int(Y),'X')
        if (sprawdz_czy_wygrana(plansza)):
            wyswietl_plansze(plansza)
            break
        plansza=tura_komputera(plansza)
        wyswietl_plansze(plansza)
    except:
        print('Nieprawidłowe dane')
        continue


