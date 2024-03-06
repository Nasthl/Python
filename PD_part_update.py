import pandas as pd

#Makro do porównania dwóch eksportów bibliotek z PD i wypisania różnic w punktach procesowych

def twin_points(df_old_points, df_new_points):
    f_df_twin_points=pd.DataFrame({'Name': df_new_points['Name']})
    f_df_twin_points['Twin name']=''
    f_df_twin_points['Delta']=''
    for i in range(len(df_new_points)):
        odleglosc=((df_new_points.iloc[i,1]-df_old_points.iloc[:,1])**2+(df_new_points.iloc[i,2]-df_old_points.iloc[:,2])**2+(df_new_points.iloc[i,3]-df_old_points.iloc[:,3])**2)**(1/2)
        odleglosc=odleglosc.dropna()
        if len(odleglosc):
            f_df_twin_points.iloc[i,1]=df_old_points.loc[odleglosc.idxmin(),'Name']
            f_df_twin_points.iloc[i,2]=odleglosc.min()
        else:
            f_df_twin_points.iloc[i, 1] = 'BRAK'
            f_df_twin_points.iloc[i, 2] = 'BRAK'
    return f_df_twin_points


df_old_input = pd.read_excel('input.xlsx', sheet_name='old', skiprows=0, usecols='A:H')
df_new_input = pd.read_excel('input.xlsx', sheet_name='new', skiprows=0, usecols='A:H')
sciezka_podsumowanie = 'PD_podsumowanie.xlsx'

# Wyodrebnienie z starej i nowej listy punktow i klejow, oraz wyciagniecie tylko kolumn name i location
#oraz podzielenie kolumny location na trzy kolumny
df_old_list = df_old_input[(df_old_input['class']=='WeldPoint') | (df_old_input['class']=='ContinuousMfg')]
df_old_list=df_old_list[['name','location']]
df_old_list['location']=df_old_list['location'].str.replace('\"','', regex=True)
df_old_list[['X','Y','Z']]=df_old_list['location'].str.split(';',expand=True)
df_old_list.drop('location',axis=1,inplace=True)
df_old_list[['X','Y','Z']]=df_old_list[['X','Y','Z']].astype(float)

df_new_list = df_new_input[(df_new_input['class']=='WeldPoint') | (df_new_input['class']=='ContinuousMfg')]
df_new_list=df_new_list[['name','location']]
df_new_list['location']=df_new_list['location'].str.replace('\"','', regex=True)
df_new_list[['X','Y','Z']]=df_new_list['location'].str.split(';',expand=True)
df_new_list[['X','Y','Z']]=df_new_list[['X','Y','Z']].astype(float)
df_new_list.drop('location',axis=1,inplace=True)

df_old_list.rename(columns={df_old_list.columns[0]: 'Name',df_old_list.columns[1]: "old_X",df_old_list.columns[2]: "old_Y",df_old_list.columns[3]: "old_Z"},inplace=True)
df_new_list.rename(columns={df_new_list.columns[0]: 'Name',df_new_list.columns[1]: "new_X",df_new_list.columns[2]: "new_Y",df_new_list.columns[3]: "new_Z"},inplace=True)

# VLookup
df_inner_join = pd.merge(df_old_list, df_new_list, on='Name', how="inner")
# Sprawdzanie przesuniecia
df_inner_join['delta'] = ((df_inner_join['old_X'] - df_inner_join['new_X']) ** 2 + (df_inner_join['old_Y'] - df_inner_join['new_Y']) ** 2 + (df_inner_join['old_Z'] - df_inner_join['new_Z']) ** 2) ** (1 / 2)
#Przesuniete pkty
df_shifted_points = df_inner_join.loc[df_inner_join['delta'] > 0]

# Usuniete pkty
l_old_points_list = set(df_old_list.iloc[:, 0]) - set(df_new_list.iloc[:, 0])
df_old_points=df_old_list.loc[df_old_list['Name'].isin(l_old_points_list)]

# Nowe pkty
l_new_points_list=set(df_new_list.iloc[:, 0]) - set(df_old_list.iloc[:, 0])
df_new_points = df_new_list.loc[df_new_list['Name'].isin(l_new_points_list)]


df_twin_points=twin_points(df_old_points,df_new_points)

with pd.ExcelWriter(sciezka_podsumowanie) as excel_podsumowanie:
    df_old_points.to_excel(excel_podsumowanie, sheet_name='OLD', index=False)
    df_new_points.to_excel(excel_podsumowanie, sheet_name='NEW', index=False)
    df_shifted_points.to_excel(excel_podsumowanie, sheet_name='SHIFTED', index=False)
    df_twin_points.to_excel(excel_podsumowanie, sheet_name='TWINS', index=False)
    #df_new_list.to_excel(excel_podsumowanie,sheet_name='Nowa_lista',index=False)