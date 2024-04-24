import pandas as pd
import json

df1 = pd.read_excel('SMpolska.xls', sheet_name='Index', dtype={
                    'Członkowie Zarządu': str})

df2 = pd.read_excel('SMpolska.xls', sheet_name='CzlonkowieZarzadu', dtype={
                    'Członkowie Zarządu': str})

numery_krs = set(df1['Członkowie Zarządu']).intersection(
    df2['Członkowie Zarządu'])

licznik_wierszy = 1
bufor = []

for numer_krs in numery_krs:
    wiersze = df2[df2['Członkowie Zarządu'] == numer_krs]

    for _, wiersz in wiersze.iterrows():
        print(f'Wiersz {licznik_wierszy}:')
        print(wiersz)
        print('\n')
        licznik_wierszy += 1

        bufor.append(wiersz.to_dict())

with open('bufor.json', 'a', encoding='utf-8') as f:
    json.dump(bufor, f, ensure_ascii=False, indent=4)
