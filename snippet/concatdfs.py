import pandas as pd
import numpy as np
from os import walk

pd.set_option('display.max_columns', 30)
pd.set_option('display.max_rows', 40)
path = 'data/GIANG'
file_names = next(walk(path), (None, None, []))[2]
file_names = [f'{path}/{x}' for x in file_names]
dfs = [pd.read_excel(x, header=1) for x in file_names]


def loc_df(df):
    df = df.loc[~pd.isna(df['Code production'])]
    return df


dfs = [loc_df(df) for df in dfs]

result = pd.concat(dfs)
result = result['Code production']

del dfs

src_ = pd.read_excel('data/(26.03.2022)_MEGAELEC_BOM ACTUAL 1.xlsx', header=8)
src_ = src_.loc[src_['status'] == 'cần tìn BOM']['code']


def convert_to_int(string):
    try:
        return str(int(string))
    except:
        return string


def look_up(string):
    words = string.split('-')
    if '-'.join(words) in src_:
        return '-'.join(words)
    else:
        try:
            words = words.pop()
            return look_up('-'.join(words))
        except:
            return None


def convert_to_format(string):
    srcs = string.split('-')
    s = '-'.join([convert_to_int(word) for word in srcs])
    return s


result = result.to_frame()
src_ = src_.to_frame()

result = result.drop_duplicates(subset=['Code production'])

result['AB'] = [convert_to_format(word) for word in result['Code production']]
result['AB'] = [look_up(word) if look_up(
    word) else word for word in result['AB']]

print('DONE')
