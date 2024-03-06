import pandas as pd
import numpy as np

#Porównuje listy punktów wraz z sprawdzeniem czy punkty dalej łączą te same blachy

# Kolumny w excelu:
# 1 - nazwa punktu
# 2-4 - koordynaty X,Y,Z
# 4-6 - nazwy partów
# 7 - nazwa robota (tylko dla starej listy)
# 8,9,... - dowolne
#skiprows - okresla ile pierwszych wierszy ma pominac w wczytywaniu

# Funkcja wyciaga liste punktów z przypisanym robotem z starej listy punktow
def our_points(df_old_list):
    return df_old_list.loc[~df_old_list['Robot'].isnull()]

 # Funkcja generuje liste blach, które są łączone za pomocą punktów z przypisanym robotem
def our_parts(df_our_points):
    return pd.unique(df_our_points[['Part_1','Part_2','Part_3']].values.ravel('K'))

def twin_points(df_old_points, df_new_points):
    f_df_twin_points=pd.DataFrame({'Name': df_new_points['Name']})
    f_df_twin_points['Twin name']=''
    f_df_twin_points['Delta']=''
    for i in range(len(df_new_points)):
        odleglosc=((df_new_points.iloc[i,1]-df_old_points.iloc[:,1])**2+(df_new_points.iloc[i,2]-df_old_points.iloc[:,2])**2+(df_new_points.iloc[i,3]-df_old_points.iloc[:,3])**2)**(1/2)
        f_df_twin_points.iloc[i,1]=df_old_points.loc[odleglosc.idxmin(),'Name']
        f_df_twin_points.iloc[i,2]=odleglosc.min()
    return f_df_twin_points


#Wczytanie danych
#df_old_list = pd.read_excel('a.xlsx', sheet_name='old', skiprows=28, usecols='G:J,Z,AK,AV,FB:FG')
#df_new_list = pd.read_excel('a.xlsx', sheet_name='new', skiprows=28, usecols='G:J,Z,AK,AV')
df_old_list = pd.read_excel('a.xlsx', sheet_name='old', skiprows=1, usecols='A:D')
df_new_list = pd.read_excel('a.xlsx', sheet_name='new', skiprows=1, usecols='A:D')


#Nazwenictwo kolumn
df_old_list.rename(columns={df_old_list.columns[0]: 'Name',df_old_list.columns[1]: "old_X",df_old_list.columns[2]: "old_Y",df_old_list.columns[3]: "old_Z"},inplace=True)
df_new_list.rename(columns={df_new_list.columns[0]: 'Name',df_new_list.columns[1]: "new_X",df_new_list.columns[2]: "new_Y",df_new_list.columns[3]: "new_Z"},inplace=True)
df_old_list.rename(columns={df_old_list.columns[4]: 'Part_1',df_old_list.columns[5]: 'Part_2',df_old_list.columns[6]: 'Part_3',df_old_list.columns[7]: 'Robot'},inplace=True)
df_new_list.rename(columns={df_new_list.columns[4]: 'Part_1',df_new_list.columns[5]: 'Part_2',df_new_list.columns[6]: 'Part_3'},inplace=True)

#DO PRZEROBIENIA

#/DO PRZEROBIENIA

#Punkty które robimy, party które zgrzewamy
#=our_points(df_old_list)
#l_our_parts_list=our_parts(df_our_points_list)

# generowanie pustego porownanie.xlsx
sciezka_podsumowanie = 'podsumowanie.xlsx'

# VLookup
df_inner_join = pd.merge(df_old_list, df_new_list, on='Name', how="inner")


df_check_parts=df_inner_join[['Name','old_X','old_Y','old_Z','Part_1_x','Material 1_x','Thickness 1_x','Part_2_x','Material 2_x','Thickness 2_x','Part_3_x','Material 3_x','Thickness 3_x','new_X','new_Y','new_Z','Part_1_y','Material 1_y','Thickness 1_y','Part_2_y','Material 2_y','Thickness 2_y','Part_3_y','Material 3_y','Thickness 3_y','AKZ','Robot']].copy()
df_check_parts.rename(columns={df_check_parts.columns[4]: 'old_Part_1',df_check_parts.columns[5]: 'old_Material_1',df_check_parts.columns[6]: 'old_Thickness_1',df_check_parts.columns[7]: 'old_Part_2',df_check_parts.columns[8]: 'old_Material_2',df_check_parts.columns[9]: 'old_Thickness_2',df_check_parts.columns[10]: 'old_Part_3',df_check_parts.columns[11]: 'old_Material_3',df_check_parts.columns[12]: 'old_Thickness_3'},inplace=True)
df_check_parts.rename(columns={df_check_parts.columns[16]: 'new_Part_1',df_check_parts.columns[17]: 'new_Material_1',df_check_parts.columns[18]: 'new_Thickness_1',df_check_parts.columns[19]: 'new_Part_2',df_check_parts.columns[20]: 'new_Material_2',df_check_parts.columns[21]: 'new_Thickness_2',df_check_parts.columns[22]: 'new_Part_3',df_check_parts.columns[23]: 'new_Material_3',df_check_parts.columns[24]: 'new_Thickness_3'},inplace=True)
df_check_parts['Check_part']=np.where((df_inner_join['Part_1_x']==df_inner_join['Part_1_y'])&(df_inner_join['Part_2_x']==df_inner_join['Part_2_y'])&((df_inner_join['Part_3_x']==df_inner_join['Part_3_y'])|(pd.isnull(df_inner_join['Part_3_x'])&pd.isnull(df_inner_join['Part_3_y']))),True,False)
df_check_parts['Check_material']=np.where((df_inner_join['Material 1_x']==df_inner_join['Material 1_y'])&(df_inner_join['Material 2_x']==df_inner_join['Material 2_y'])&((df_inner_join['Material 3_x']==df_inner_join['Material 3_y'])|(pd.isnull(df_inner_join['Material 3_x'])&pd.isnull(df_inner_join['Material 3_y']))),True,False)
df_check_parts['Check_thickness']=np.where((df_inner_join['Thickness 1_x']==df_inner_join['Thickness 1_y'])&(df_inner_join['Thickness 2_x']==df_inner_join['Thickness 2_y'])&((df_inner_join['Thickness 3_x']==df_inner_join['Thickness 3_y'])|(pd.isnull(df_inner_join['Thickness 3_x'])&pd.isnull(df_inner_join['Thickness 3_y']))),True,False)


# Sprawdzanie przesuniecia
df_inner_join['delta'] = ((df_inner_join['old_X'] - df_inner_join['new_X']) ** 2 + (df_inner_join['old_Y'] - df_inner_join['new_Y']) ** 2 + (df_inner_join['old_Z'] - df_inner_join['new_Z']) ** 2) ** (1 / 2)
#Przesuniete pkty
df_shifted_points = df_inner_join.loc[df_inner_join['delta'] > 0]
#df_our_shifted_points=df_shifted_points[df_shifted_points['Robot'].notnull()]

# Usuniete pkty
l_old_points_list = set(df_old_list.iloc[:, 0]) - set(df_new_list.iloc[:, 0])
df_old_points=df_old_list.loc[df_old_list['Name'].isin(l_old_points_list)]
#df_our_old_points=df_old_points[df_old_points['Robot'].notnull()]

# Nowe pkty
l_new_points_list=set(df_new_list.iloc[:, 0]) - set(df_old_list.iloc[:, 0])
df_new_points = df_new_list.loc[df_new_list['Name'].isin(l_new_points_list)]

# Nowe pkty z blachami, ktore zgrzewamy
# Co jezeli jest ta sama nazwa pktu ale inne blachy?
#df_our_new_points=df_new_points.loc[df_new_points['Part_1'].isin(l_our_parts_list) & df_new_points['Part_2'].isin(l_our_parts_list) & df_new_points['Part_3'].isin(l_our_parts_list)]

df_twin_points=twin_points(df_old_points,df_new_points)
#print(df_twin_points)

#Zapisywanie arkuszy
with pd.ExcelWriter(sciezka_podsumowanie) as excel_podsumowanie:
    #df_our_new_points.to_excel(excel_podsumowanie,sheet_name='OUR_NEW_POINTS',index=False)
   # df_our_old_points.to_excel(excel_podsumowanie,sheet_name='OUR_OLD_POINTS',index=False)
   # df_our_shifted_points.to_excel(excel_podsumowanie,sheet_name='OUR_SHITED_POINTS',index=False)
    df_old_points.to_excel(excel_podsumowanie,sheet_name='OLD',index=False)
    df_shifted_points.to_excel(excel_podsumowanie, sheet_name='SHIFTED', index=False)
    df_new_points.to_excel(excel_podsumowanie, sheet_name='NEW', index=False)
    df_inner_join.to_excel(excel_podsumowanie,sheet_name='INNER', index=False)
    df_twin_points.to_excel(excel_podsumowanie,sheet_name='TWINS', index=False)
    #df_check_parts.to_excel(excel_podsumowanie,sheet_name='Parts_check',index=False)


#TODO
#1. sieci do zaproponowania  robotow na nowe pkty
#4. wpisanie do MPL listy :
#4a. wpisanie nieprzesunietych pktow
#4b. wpisanie przesunietych pktow i zaznaczenie ich
#4c. zaznaczenie naszych nowych pktow

