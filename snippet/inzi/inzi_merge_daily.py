import warnings

import pandas as pd

pd.options.mode.chained_assignment = None
pd.set_option('display.max_columns', 80)
pd.set_option('display.max_rows', 55)

duplicated_item_file = '../../output/duplicated_item.txt'
goods_processing_file = '../../data/inzi/(JUNE)Following goods processing.xlsx'
previous_file = '../../data/inzi/02-11-2021_IN-OUT-STOCK_DAILY REPORT.xlsb'
today_file = '../../data/inzi/02-11-2021_IN-OUT DAILY REPORT.xlsb'
open_stock = previous_file
out_put_file = '../../output/dailyconcat.xlsx'

f = open(duplicated_item_file, 'w')

goods_processing = pd.read_excel(goods_processing_file, sheet_name=None)

stock = pd.read_excel(open_stock, header=2)

stock.rename(columns={
    'Unnamed: 5': 'Material'
}, inplace=True)
stock['OPENING STOCK'] = stock['Unrestricted'] + \
    stock['Inspection'] + stock['Returned'] + stock['Blocked']
stock = stock[['Material', 'OPENING STOCK']]
stock = stock.groupby(stock['Material']).aggregate(
    'sum')


def get_result(sheetname):
    previous = pd.read_excel(
        previous_file, sheet_name=sheetname, header=10)
    today = pd.read_excel(
        today_file, sheet_name=sheetname, header=8)

    previous.rename(columns={
        'Unnamed: 1': 'Material',
        'Unnamed: 2': 'UOM'
    }, inplace=True)
    today.rename(columns={
        'Unnamed: 1': 'Material',
        'Unnamed: 2': 'UOM'
    }, inplace=True)

    del previous['Unnamed: 0']
    del today['Unnamed: 0']
    del today['IN']
    del today['OUT']

    previous.columns = [*previous.columns[:-1], 'CLOSING STOCK']

    result = pd.concat([previous, today])

    result = result[['Material', 'UOM']]

    del previous['UOM']
    del today['UOM']

    result.drop_duplicates(inplace=True, keep='first')

    duplicate = result[result.duplicated(subset=['Material'])]
    duplicate.to_excel('../../output/duplicated.xlsx')

    result = pd.merge(result, stock, how="left", on=['Material'], sort=False,
                      indicator=False, validate=None)
    result = pd.merge(result, previous, how="left", on=['Material'], sort=False,
                      indicator=False, validate=None)
    result = pd.merge(result, today, how="left", on=['Material'], sort=False,
                      indicator=False, validate=None)

    return result


print('Start merging...')

raw_concat = get_result('REPORT-RAW')
wip_concat = get_result('REPORT-WIP')
fg_concat = get_result('REPORT-FG')

raw_concat.to_excel('../../output/fixing')

print('Done')

print('Start checking duplicated items...')

check_case_one = pd.merge(raw_concat, fg_concat, how="inner", on=['Material'], sort=False,
                          indicator=False, validate=None)


check_case_one_items = check_case_one["Material"].unique().tolist()
f.write(f'{len(check_case_one_items)} items present in RAW and FG\n')
f.write('START CHECKING --------------\n')
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
            f.write(
                f'Item {material} in trading of {name} ')

        elif material in sheet['Processing'].unique():
            in_processing = True
            f.write(
                f'Item {material} in processing of {name} ')
        else:
            pass

    if in_trading and in_processing:
        f.write(
            f'** Item {material} in trading and processing as well, please check **\n')
    elif in_trading and not in_processing:
        f.write(
            f'** Item {material} in trading only --> Should move the item to RAW and remove from FG **\n')
        # raw_concat = raw_concat.append(
        #     fg_concat[fg_concat['Material'] == material])
        # fg_concat = fg_concat.drop(
        #     fg_concat[fg_concat['Material'] == material].index)
        check_case_one_items.remove(material)
    elif not in_trading and in_processing:
        f.write(
            f'** Item {material} in processing only --> Should move the item to FG and remove from RAW **\n')
        # fg_concat = fg_concat.append(
        #     raw_concat[raw_concat['Material'] == material])
        # raw_concat = raw_concat.drop(
        #     raw_concat[raw_concat['Material'] == material].index)
        check_case_one_items.remove(material)


if check_case_one_items:
    f.write(
        f'There is {check_case_one_items} can not be define in RAW or FG, please check again...\n')


f.write('---------------------------------------------------------------------\n')

''' CHECK CASE TWO '''
check_case_two = pd.merge(raw_concat, wip_concat, how="inner", on=['Material'], sort=False,
                          indicator=False, validate=None)


check_case_two_items = check_case_two["Material"].unique().tolist()
f.write(f'{len(check_case_two_items)} items present in RAW and WIP\n')
f.write('START CHECKING--------\n')

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
            f.write(
                f'Item {material} in trading of {name} ')

        elif material in sheet['Processing'].unique():
            in_processing = True
            f.write(
                f'Item {material} in processing of {name} ')
        else:
            pass

    if in_trading and in_processing:
        f.write(
            f'** Item {material} in trading and processing as well, please check **\n')
    elif in_trading and not in_processing:
        f.write(
            f'** Item {material} in trading only --> Should move the item to RAW and remove from WIP **\n')
        # raw_concat = raw_concat.append(
        #     wip_concat[wip_concat['Material'] == material])
        # wip_concat = wip_concat.drop(
        #     wip_concat[wip_concat['Material'] == material].index)
        check_case_two_items.remove(material)
    elif not in_trading and in_processing:
        f.write(
            f'** Item {material} in processing only --> Should move the item to WIP and remove from RAW **\n')
        # wip_concat = wip_concat.append(
        #     raw_concat[raw_concat['Material'] == material])
        # raw_concat = raw_concat.drop(
        #     raw_concat[raw_concat['Material'] == material].index)
        check_case_two_items.remove(material)

if check_case_two_items:
    f.write(
        f'There is {check_case_two_items} can not be define RAW or WIP, please check...\n')

f.close()

print('Done')

with pd.ExcelWriter(out_put_file) as writer:
    raw_concat.to_excel(writer, sheet_name='RAW-REPORT')
    wip_concat.to_excel(writer, sheet_name='WIP-REPORT')
    fg_concat.to_excel(writer, sheet_name='FG-REPORT')

print(
    f'The output is {out_put_file} and duplicated items is in {duplicated_item_file}')
