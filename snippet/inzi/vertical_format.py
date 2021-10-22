import datetime
import pandas as pd
import numpy as np

pd.options.mode.chained_assignment = None
pd.set_option('display.max_columns', 150)
pd.set_option('display.max_rows', 100)

file_url = '../../data/inzi/03-10-2021_IN-OUT-STOCK_DAILY REPORT.xlsb'
output_url = '../../output/verti.xlsx'

print('Running................')
stock = pd.read_excel(file_url, sheet_name='REPORT-RAW', header=9)
get_date = pd.read_excel(file_url, sheet_name='REPORT-RAW', header=7)
stock.rename(columns={
    'Unnamed: 1': 'CODE',
    'Unnamed: 2': 'UOM',
    'Unnamed: 3': 'OPENING STOCK',
}, inplace=True)
start_date = None
for x in get_date.columns:
    if isinstance(x, int):
        start_date = x
        break


date = datetime.date(1900, 1, 1) + datetime.timedelta(start_date - 2)

index_to_indice = []
for index, column in enumerate(stock.columns):
    if column.startswith('PURCHASE'):
        index_to_indice.append(index)


numbers_of_days = len(index_to_indice)
list_df = [None] * numbers_of_days
for no, index in enumerate(index_to_indice):
    if no == 0:
        list_df[no] = stock.iloc[:, list(
            range(1, 4)) + list(range(index, index_to_indice[no + 1]))]
    elif no < numbers_of_days - 1:
        list_df[no] = stock.iloc[:, list(
            range(1, 4)) + list(range(index-1, index_to_indice[no + 1]))]
    else:
        list_df[no] = stock.iloc[:, list(
            range(1, 4)) + list(range(index-1, len(stock.columns)))]

for no, df in enumerate(list_df):
    df.columns = [*df.columns[:-1], 'CLOSING STOCK']
    df['DATE'] = date + datetime.timedelta(no)
    df.rename(columns=lambda x: x.split('.')[0], inplace=True)
    if no > 0:
        del df['OPENING STOCK']
        df.columns.values[2] = 'OPENING STOCK'
    columns = []
    for column in df.columns:
        if column not in columns:
            columns.append(column)
        else:
            columns.append(column + '_new')
    df.columns = columns


result = pd.concat(list_df)

columns = result.columns.to_list()
# columns.append(columns.pop(columns.index('CLOSING STOCK')))
columns.pop(columns.index('DATE'))
columns.insert(2, 'DATE')

result = result[columns]
result.to_excel(output_url)
print('Done.')
