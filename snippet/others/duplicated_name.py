from functools import reduce
import pandas as pd
import datetime

pd.set_option('display.max_columns', 45)
pd.set_option('display.max_rows', 100)


begin = datetime.datetime.now()

excel = pd.read_excel('data/5402.xlsx')


# Format
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.translate(
    {ord(i): '' for i in ',;'})
excel['Don vi doi tac'] = excel['Don vi doi tac'] + ' '

excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    '( ', '(', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' )', ')', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ')', ') ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    '(', ' (', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    '-', ' - ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    '&', ' & ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' .LTD', ' LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' CO.LTD', ' CO. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' CO LTD', ' CO. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' CO.LTD.', ' CO. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' CO . LTD', ' CO. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' CO. LTD.', ' CO. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' CO. LTD .', ' CO. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' PVT.LTD', ' PVT. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' LTD.', ' LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' PTE', ' PTE.', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' G.A.', ' G.A', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' FTY.LTD', ' FTY. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' FTY LTD', ' FTY. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' IND. ', ' IND ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' KUN SHAN ', ' KUNSHAN ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' AND ', ' & ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' SDN BHD', ' SDN. BHD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' SDN.BHD', ' SDN. BHD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' SDN. BHD.', ' SDN. BHD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' SDN.BHD.', ' SDN. BHD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' SDN. BHD .', ' SDN. BHD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' EXP.CO.LTD', ' EXP CO. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' COLTD', ' CO. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' PTE LTD', ' PTE. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' PTE.LTD', ' PTE. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' MFG.CO.LTD', ' MFG. CO. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' MFG. CO.LTD', ' MFG. CO. LTD', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' PETROCHEMICALS.', ' PETROCHEMICALS', regex=False)

# Spelling Errors
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' LIMTED ', ' LIMITED ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' SYNTHETICES ', ' SYNTHETICS ')
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' TRAADING ', ' TRADING ')
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' COPORATION ', ' CORPORATION ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    'SHANG HAI', 'SHANGHAI', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' EXPORT', ' EXPORT ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' FIRNITURE', ' FURNITURE', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' MATERIAL ', ' FURNITURE', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' MATERIALSS ', ' MATERIALS ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' INT L ', " INT'L ", regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    'ZHE JIANG', 'ZHEJIANG', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' INDODRAMA ', ' INDORAMA ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' FIBRE ', ' FIBER ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' CORPERATION ', ' CORPORATION ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' INTERNATIONALHOLDINGS ', ' INTERNATIONAL HOLDINGS ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' TEXTTILE ', ' TEXTILE ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' TEXILE ', ' TEXTILE ', regex=False)
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' FEBERS ', ' FIBERS ', regex=False)

# Remove multiple space with one space

excel['Don vi doi tac'] = [' '.join(x.split())
                           for x in excel['Don vi doi tac']]
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.rstrip('.')

# For some reason, it's here...
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    'VIET NAM', 'VIETNAM')
excel['Don vi doi tac'] = excel['Don vi doi tac'].str.replace(
    ' VIETNAM', ' (VIETNAM)')


words_to_remove = [
    ' CO.',
    'CONG TY',
    ' CO PHAN',
    ' TNHH',
    'CTY',
    ' PTE.',
    ' LTD',
    ' LIMITED',
    ' INC',
    ' CORPORATION',
    " INT'L",
    ' COMPANY',
    ' FTY'
]

excel['TEMP'] = reduce(lambda a, b: a.str.replace(
    b, '', regex=False), [excel['Don vi doi tac']] + words_to_remove)
excel['TEMP'] = excel['TEMP'].str.rstrip('.')
excel['TEMP'] = [' '.join(x.split()) for x in excel['TEMP']]

excel_count = len(excel.index)
for index, row in excel.iterrows():
    if index < excel_count - 1:
        current_cell = (excel['TEMP'].iloc[index].split())
        next_cell = (excel['TEMP'].iloc[index + 1].split())
        if len(set(current_cell).intersection(next_cell)) / len(current_cell) >= 0.5 or len(set(current_cell).intersection(next_cell)) / len(next_cell) >= 0.5:
            excel.at[index + 1, 'TEMP'] = excel.at[index, 'TEMP']


def final_col(row):
    if set(row['TEMP'].split()) <= set(row['TEMP'].split()):
        return row['TEMP']


excel.apply(lambda row: final_col(row), axis=1)

excel.to_excel('output/testing.xlsx')
end = datetime.datetime.now()

print(f'{end - begin}s to run')
