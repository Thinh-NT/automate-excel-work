import pandas as pd
import numpy as np
from os import walk

pd.set_option('display.max_columns', 30)
pd.set_option('display.max_rows', 40)
path = '../../data/BOM T6'
file_names = next(walk(path), (None, None, []))[2]
file_names = (f'{path}/{x}' for x in file_names)
dfs = (pd.read_excel(x, header=1) for x in file_names)
bom = pd.concat(dfs)
bom.to_excel('../../output/fuck.xlsx')
del bom['Unnamed: 0']

# Create level
bom['Level'] = bom['Level'].astype(str)
bom['Level'] = bom['Level'].str[-1]
try:
    bom['Level'] = pd.to_numeric(bom['Level'])
except ValueError:
    print('Error')
    quit()


def find_if_item_is_child(row):
    if row['Level'] > 0:
        return True


bom['is_child'] = bom.apply(lambda row: find_if_item_is_child(row), axis=1)


def create_columns(row, i):
    if row['Level'] == i:
        return row['Component']


for i in bom['Level'].unique():
    bom[i] = bom.apply(lambda row: create_columns(row, i), axis=1)

bom.ffill(inplace=True)


def fill_father_name(row):
    if row['is_child'] and row['Level'] > 0:
        return row[row['Level'] - 1]


bom['Father'] = bom.apply(lambda row: fill_father_name(row), axis=1)
bom['Product'] = bom[0]
del bom['is_child']
for i in bom['Level'].unique():
    del bom[i]

cols = ['Level', 'Component', 'Product',
        'Father', 'Deleted Item', 'Description', 'Specification',
        'Unit', 'Quantity', 'Stock', 'Unnamed: 9', 'Process', 'Provision',
        'Bulk', 'Material Type', 'MRP Type', 'Procurement',
        'Special Procurement', 'Component Scrap', 'Operation Scrap']
bom = bom[cols]

bom.to_excel('../../output/finish.xlsx')
print('Done')
