# import numpy as np
# from functools import reduce
import warnings

import pandas as pd

pd.options.mode.chained_assignment = None

warnings.simplefilter(action='ignore', category=FutureWarning)


open_stock = '../data/inzi/stock.xlsb'
goods_processing_file = '../data/inzi/goods processing.xlsx'
files = [
    '../data/inzi/02 SEP/01-09-2021_IN-OUT DAILY REPORT.xlsb',
    '../data/inzi/03 SEP/02-09-2021_IN-OUT DAILY REPORT.xlsb',
    '../data/inzi/04 SEP/03-09-2021_IN-OUT DAILY REPORT.xlsb',
]
out_put_file = '../output/concet.xlsx'


stock = pd.read_excel(open_stock, header=2)
goods_processing = pd.read_excel(
    goods_processing_file, sheet_name=None)

stock.rename(columns={
    'Unnamed: 5': 'Materials'
}, inplace=True)
stock['OPENING STOCK'] = stock['Unrestricted'] + \
    stock['Inspection'] + stock['Returned'] + stock['Blocked']
stock = stock[['Materials', 'OPENING STOCK']]
stock.rename(columns={'Materials': 'Material'}, inplace=True)
stock = stock.groupby(stock['Material']).aggregate(
    'sum')


def get_concat(sheet):
    dfs = []
    for file_name in files:
        df = pd.read_excel(file_name, sheet_name=sheet, header=8)
        df.drop(columns=['Unnamed: 0', 'Unnamed: 2',
                         'IN', 'OUT'], inplace=True)
        df.rename(columns={'Unnamed: 1': 'Material'}, inplace=True)
        if 'Unnamed: 5' in df:
            df.rename(columns={'Unnamed: 5': 'PURCHASE'}, inplace=True)
        dfs.append(df)

    final_result = pd.concat(dfs)
    final_result['temp'] = 0
    final_result = final_result[['Material', 'temp']]
    final_result.drop_duplicates(inplace=True)
    final_result = pd.merge(final_result, stock, how="left", on=['Material'], sort=False,
                            indicator=False, validate=None)
    del final_result['temp']
    l = len(dfs)

    for _ in range(l):
        in_stock = [344, 401, 623, 720, 321, 323, 327, 602]
        sum_col = []
        out_stock = [401, 720]
        subtract_col = []

        dfs[0] = pd.merge(dfs[0], final_result[final_result.columns[[0, -1]]], how="outer", on=['Material'], sort=False,
                          indicator=False, validate=None)
        dfs[0].set_axis([*dfs[0].columns[:-1], 'temporary'],
                        axis=1, inplace=True)

        for column in in_stock:
            if column in dfs[0].columns:
                sum_col.append(column)
        sum_col.append('temporary')

        for column in out_stock:
            if column in dfs[0].columns:
                subtract_col.append(column)

        dfs[0]['CLOSING STOCK'] = dfs[0][sum_col].sum(
            axis=1) - dfs[0][subtract_col].sum(axis=1)
        del dfs[0]['temporary']

        final_result = pd.merge(final_result, dfs[0], how="outer", on=['Material'], sort=False,
                                indicator=False, validate=None)
        del dfs[0]
    print(f'Concat {l} files together in {sheet}')
    # final_result.drop(final_result.columns[-1], axis=1, inplace=True)
    final_result.fillna(0, inplace=True)
    return final_result


raw_concat = get_concat('REPORT-RAW')
wip_concat = get_concat('REPORT-WIP')
fg_concat = get_concat('REPORT-FG')
print('---------------------------------------------------------------------')


class Object(object):
    pass


''' CHECK CASE ONE '''


check_case_one = pd.merge(raw_concat, fg_concat, how="inner", on=['Material'], sort=False,
                          indicator=False, validate=None)


check_case_one_items = check_case_one["Material"].unique().tolist()
print(f'{len(check_case_one_items)} items present in RAW and FG')
print('START CHECKING --------------')
for index, row in check_case_one.iterrows():
    material = row['Material']
    in_trading = False
    in_processing = False
    for name, sheet in goods_processing.items():
        sheet = sheet[['Unnamed: 1', 'Unnamed: 12']]
        sheet.rename(columns={
            'Unnamed: 1': 'Trading', 'Unnamed: 12': 'Processing',
        }, inplace=True)

        if material in sheet['Trading'].unique():
            in_trading = True
            print(
                f'Item {material} in trading of {name}')

        elif material in sheet['Processing'].unique():
            in_processing = True
            print(
                f'Item {material} in processing of {name}')
        else:
            pass

    if in_trading and in_processing:
        print(
            f'** Item {material} in trading and processing as well, please check **')
    elif in_trading and not in_processing:
        print(
            f'**Item {material} in trading only --> Moved the item to RAW and remove from FG**')
        raw_concat = raw_concat.append(
            fg_concat[fg_concat['Material'] == material])
        fg_concat = fg_concat.drop(
            fg_concat[fg_concat['Material'] == material].index)
        check_case_one_items.remove(material)
    elif not in_trading and in_processing:
        print(
            f'**Item {material} in processing only --> Moved the item to FG and remove from RAW**')
        fg_concat = fg_concat.append(
            raw_concat[raw_concat['Material'] == material])
        raw_concat = raw_concat.drop(
            raw_concat[raw_concat['Material'] == material].index)
        check_case_one_items.remove(material)
    print('------------')


if check_case_one_items:
    print(
        f'There is {check_case_one_items} can not be define in RAW or FG, please check again...')


print('---------------------------------------------------------------------')

''' CHECK CASE TWO '''
check_case_two = pd.merge(raw_concat, wip_concat, how="inner", on=['Material'], sort=False,
                          indicator=False, validate=None)


check_case_two_items = check_case_two["Material"].unique().tolist()
print(f'{len(check_case_two_items)} items present in RAW and WIP')
print('START CHECKING--------')

for index, row in check_case_two.iterrows():
    material = row['Material']
    in_trading = False
    in_processing = False
    for name, sheet in goods_processing.items():
        sheet = sheet[['Unnamed: 1', 'Unnamed: 12']]
        sheet.rename(columns={
            'Unnamed: 1': 'Trading', 'Unnamed: 12': 'Processing',
        }, inplace=True)

        if material in sheet['Trading'].unique():
            in_trading = True
            print(
                f'Item {material} in trading of {name}')

        elif material in sheet['Processing'].unique():
            in_processing = True
            print(
                f'Item {material} in processing of {name}')
        else:
            pass

    if in_trading and in_processing:
        print(
            f'** Item {material} in trading and processing as well, please check **')
    elif in_trading and not in_processing:
        print(
            f'**Item {material} in trading only --> Moved the item to RAW and remove from WIP**')
        raw_concat = raw_concat.append(
            wip_concat[wip_concat['Material'] == material])
        wip_concat = wip_concat.drop(
            wip_concat[wip_concat['Material'] == material].index)
        check_case_two_items.remove(material)
    elif not in_trading and in_processing:
        print(
            f'**Item {material} in processing only --> Moved the item to WIP and remove from RAW**')
        wip_concat = wip_concat.append(
            raw_concat[raw_concat['Material'] == material])
        raw_concat = raw_concat.drop(
            raw_concat[raw_concat['Material'] == material].index)
        check_case_two_items.remove(material)
    print('-------------')

if check_case_two_items:
    print(
        f'There is {check_case_two_items} can not be define RAW or WIP, please check...')


def get_df_name(df):
    name = [x for x in globals() if globals()[x] is df][0]
    return name


with pd.ExcelWriter(out_put_file) as writer:
    raw_concat = raw_concat.groupby(raw_concat['Material']).aggregate('sum')
    raw_concat.columns = [str(x)[0:3] if str(x)[0:3].isdigit()
                          else str(x).split('_')[0] for x in raw_concat.columns]
    raw_concat.to_excel(writer, sheet_name='RAW-REPORT')

    wip_concat = wip_concat.groupby(wip_concat['Material']).aggregate('sum')
    wip_concat.columns = [str(x)[0:3] if str(x)[0:3].isdigit()
                          else str(x).split('_')[0] for x in wip_concat.columns]
    wip_concat.to_excel(writer, sheet_name='WIP-REPORT')

    fg_concat = fg_concat.groupby(fg_concat['Material']).aggregate('sum')
    fg_concat.columns = [str(x)[0:3] if str(x)[0:3].isdigit()
                         else str(x).split('_')[0] for x in fg_concat.columns]
    fg_concat.to_excel(writer, sheet_name='FG-REPORT')


# dfs = []

# suffixes=(f'_{i:02}', f'_{(i + 1):02}',)


# def merge_result(left, right):

#     in_stock = [344, 401, 623, 720, 321, 323, 327, 602]
#     sum_col = []
#     out_stock = [401, 720]
#     subtract_col = []

#     right = pd.merge(right, left[left.columns[[0, -1]]], how="outer", on=['Material'], sort=False,
#                      indicator=False, validate=None)
#     right.set_axis([*right.columns[:-1], 'temporary'], axis=1, inplace=True)

#     for column in in_stock:
#         if column in right.columns:
#             sum_col.append(column)
#     sum_col.append('temporary')

#     for column in out_stock:
#         if column in right.columns:
#             subtract_col.append(column)

#     right['in_stock'] = right[sum_col].sum(
#         axis=1) - right[subtract_col].sum(axis=1)
#     del right['temporary']

#     result = pd.merge(left, right, how="outer", on=['Material'], sort=False,
#                       indicator=False, validate=None)
#     del left
#     del right
#     result.drop_duplicates(keep='first', inplace=True)
#     return result


# for file_name in files:
#     df = pd.read_excel(file_name)
#     df = df.drop(columns=['Unnamed: 0', 'IN', 'OUT'])
#     dfs.append(df)


# final_result = pd.concat(dfs)
# final_result['temp'] = 0
# final_result = final_result[['Material', 'temp']]
# final_result.drop_duplicates(inplace=True)
# final_result = pd.merge(final_result, stock, how="left", on=['Material'], sort=False,
#                         indicator=False, validate=None)

# for counter in range(l):
#     in_stock = [344, 401, 623, 720, 321, 323, 327, 602]
#     sum_col = []
#     out_stock = [401, 720]
#     subtract_col = []

#     dfs[0] = pd.merge(dfs[0], final_result[final_result.columns[[0, -1]]], how="outer", on=['Material'], sort=False,
#                       indicator=False, validate=None)
#     dfs[0].set_axis([*dfs[0].columns[:-1], 'temporary'], axis=1, inplace=True)

#     for column in in_stock:
#         if column in dfs[0].columns:
#             sum_col.append(column)
#     sum_col.append('temporary')

#     for column in out_stock:
#         if column in dfs[0].columns:
#             subtract_col.append(column)

#     dfs[0]['in_stock'] = dfs[0][sum_col].sum(
#         axis=1) - dfs[0][subtract_col].sum(axis=1)
#     del dfs[0]['temporary']

#     final_result = pd.merge(final_result, dfs[0], how="outer", on=['Material'], sort=False,
#                             indicator=False, validate=None)
#     print(f'Concat {counter + 1} files')
#     del dfs[0]
