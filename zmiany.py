import pandas as pd

# Wczytaj pierwszy arkusz do DataFrame, traktując kolumnę z numerami KRS jako ciągi znaków
df1 = pd.read_excel('SMpolska.xls', sheet_name='Index', dtype={
                    'Członkowie Zarządu': str})

# Wczytaj drugi arkusz do innego DataFrame
df2 = pd.read_excel('SMpolska.xls', sheet_name='CzlonkowieZarzadu', dtype={
                    'Członkowie Zarządu': str})

# Znajdź numery KRS, które są obecne w obu arkuszach
numery_krs = set(df1['Członkowie Zarządu']).intersection(
    df2['Członkowie Zarządu'])

# Inicjalizuj licznik wierszy
licznik_wierszy = 1

# Iteruj przez numery KRS
for numer_krs in numery_krs:
    # Znajdź wiersze w df2, które zawierają ten numer KRS
    wiersze = df2[df2['Członkowie Zarządu'] == numer_krs]

    # Wydrukuj te wiersze z numeracją
    for _, wiersz in wiersze.iterrows():
        print(f'Wiersz {licznik_wierszy}:')
        print(wiersz)
        print('\n')
        licznik_wierszy += 1
