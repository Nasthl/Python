import datetime

import pandas as pd
# Makro do wypisywania osób na projekcie wraz z terminami do kiedy sa na podstawie Cplan

def f_raport_cplan(nr,df_plan,date):
    #wywal dni przed dzis
    df_plan.drop(df_plan.columns[3:df_plan.columns.get_loc(date)-1], axis=1, inplace=True)
    #jesli dzis jest swieto to date to nastepny pracujacy dzien
    # NIE DZIALA
    #while df_plan[date].isna().all():
    #   print(f'Data to {df_plan[date]}')
    #   date = date + datetime.timedelta(days=1)
    # wywal dni swiąteczne
    df_plan=df_plan.dropna(axis=1, how='all')
    #lista wszystkich nieobecnych
    df_plan_absence=df_plan.loc[df_plan[date].isin(['U','K','E','S','Tr','Z'])]
    #Usun wszystkich co maja nieobecnosc i nie pracuja na projekcie
    for i in range(len(df_plan_absence)):
        if not (df_plan_absence.iloc[i,:].isin([nr]).any()):
            df_plan.drop(df_plan[df_plan['Name']==df_plan_absence.iloc[i,1]].index, inplace=True)
    #Wyczyszczona lista
    for i in range (len(df_plan)):
        for j in enumerate(df_plan.columns):
            if j[0]>3 and not(df_plan.iloc[i,j[0]] in ['U','K','E','S','Tr','Z',nr]) or j[0]==len(df_plan.columns)-1:
                break;
            temp=j[1]
        k=3
        urlop = False
        print(f'{df_plan.iloc[i, 1]} do ', temp.strftime('%Y/%m/%d'),end="    " )
        while k<=j[0]:
            if df_plan.iloc[i,k] in ['U','K','E','S','Tr','Z'] and not urlop:
                zacznij_urlop=df_plan.columns[k].strftime('%Y/%m/%d')
                urlop=True
            if (df_plan.iloc[i,k] == nr ) and urlop:
                skoncz_urlop=df_plan.columns[k-1].strftime('%Y/%m/%d')
                urlop=False
                print(f'Urlop {zacznij_urlop}-{skoncz_urlop}', end=" ")
            k=k+1
        if urlop:
            skoncz_urlop = df_plan.columns[k - 2].strftime('%Y/%m/%d')
            print(f'Urlop {zacznij_urlop}-{skoncz_urlop}', end=" ")
        print()
    return df_plan

def zlicz_dni(nr,df,poczatek,koniec):
    dni=0;
    #print(df.columns[poczatek],df.columns[koniec])
    for col in df.columns[poczatek:koniec+1]:
        if nr in df[col].value_counts().index:
            dni = dni + df[col].value_counts()[nr]
       # print(dni)
    return dni


today=datetime.date.today()
#today=datetime.date(2023,10,31)

print(today)
Cplan_path='C:\OneDrive\OneDrive - Chropynska\Cplan_backup\Cplan_'+str(today.strftime('%Y%m%d'))+'.xlsm'
Cplan_summary_path='C:\Porsche PO546\Cplan_check\\Cplan_check.xlsx'
nr_projektu='001'
sim_price=32
cc_price=30
pln_price=32
miesiac=today.strftime('%m')
rok=today.strftime('%Y')
s_miesiac=datetime.date(int(rok),int(miesiac),1)
if int(miesiac)==12:
    miesiac=1
    rok=int(rok)+1
else:
    miesiac=int(miesiac)+1
e_miesiac=datetime.date(int(rok),int(miesiac),1)
e_miesiac=e_miesiac-datetime.timedelta(days=1)
df_cplan=pd.read_excel(Cplan_path,'Kappaplan',skiprows=2)
df_cplan.rename(columns={"C-Plan":"PersonalNr","Unnamed: 1":"Name","Unnamed: 2":"Kst"},inplace=True)
df_cplan_eng=df_cplan.loc[df_cplan['Kst'].isin(['SIM','CC','PLN'])]
df_cplan_cc=df_cplan.loc[df_cplan['Kst'].isin(['CC'])]
df_cplan_pln=df_cplan.loc[df_cplan['Kst'].isin(['PLN'])]
df_cplan_sim=df_cplan.loc[df_cplan['Kst'].isin(['SIM'])]
s_miesiac_index=df_cplan.columns.get_loc(datetime.datetime.combine(s_miesiac,datetime.time(0,0,0)))
e_miesiac_index=df_cplan.columns.get_loc(datetime.datetime.combine(e_miesiac,datetime.time(0,0,0)))
today_temp=datetime.datetime.combine(today,datetime.time(0,0,0))
today_index=df_cplan.columns.get_loc(today_temp)
dni_sim=zlicz_dni(nr_projektu,df_cplan_sim,s_miesiac_index,e_miesiac_index)
dni_pln=zlicz_dni(nr_projektu,df_cplan_pln,s_miesiac_index,e_miesiac_index)
dni_cc=zlicz_dni(nr_projektu,df_cplan_cc,s_miesiac_index,e_miesiac_index)
przepalone_dni_sim=zlicz_dni(nr_projektu,df_cplan_sim,s_miesiac_index,today_index)
przepalone_dni_pln=zlicz_dni(nr_projektu,df_cplan_pln,s_miesiac_index,today_index)
przepalone_dni_cc=zlicz_dni(nr_projektu,df_cplan_cc,s_miesiac_index,today_index)
print(f'dni sim {dni_sim} dni pln {dni_pln} dni cc {dni_cc}')
print(f'przepalone dni sim {przepalone_dni_sim} dni pln {przepalone_dni_pln} dni cc {przepalone_dni_cc}')
print(f'Szacowany całkowity koszt {dni_sim*sim_price*8+dni_cc*cc_price*8+dni_pln*pln_price*8} euro')
#print(f'Szacowany całkowity koszt SIM {dni_sim*8} {dni_sim*sim_price*8} euro')
#print(f'Szacowany całkowity koszt CC {dni_cc*8} {dni_cc*cc_price*8} euro')
#print(f'Szacowany całkowity koszt PLA {dni_pln*8} {dni_pln*pln_price*8} euro')
print(f'przepalone godziny {przepalone_dni_sim*sim_price*8+przepalone_dni_cc*cc_price*8+przepalone_dni_pln*pln_price*8} euro')


df_cplan_eng_project=df_cplan_eng.loc[df_cplan_eng[today_temp].isin([nr_projektu,'U','K','E','S','Tr','Z'])]
print('Tu')
df_cplan_eng_project_summary=f_raport_cplan(nr_projektu,df_cplan_eng_project,today_temp)
print('a tu?')
with pd.ExcelWriter(Cplan_summary_path) as excel:
    df_cplan_eng_project_summary.to_excel(excel, sheet_name=str(today.strftime('%Y%m%d')), index=False)