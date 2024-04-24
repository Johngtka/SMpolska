import pandas as pd
import json

buffer = []
rowCounter = 1

indexKRS = pd.read_excel('SMpolska.xls', sheet_name='Index', dtype={
    'Członkowie Zarządu': str})

boardMembersKRS = pd.read_excel('SMpolska.xls', sheet_name='CzlonkowieZarzadu', dtype={
    'Członkowie Zarządu': str})

relationKRS = set(indexKRS['Członkowie Zarządu']).intersection(
    boardMembersKRS['Członkowie Zarządu'])

for numberKRS in relationKRS:
    rows = boardMembersKRS[boardMembersKRS['Członkowie Zarządu'] == numberKRS]

    for _, row in rows.iterrows():
        print(f'Wiersz {rowCounter}:')
        print(row)
        print('\n')
        rowCounter += 1
        buffer.append(row.to_dict())

with open('bufor.json', 'a', encoding='utf-8') as f:
    json.dump(buffer, f, ensure_ascii=False, indent=4)
