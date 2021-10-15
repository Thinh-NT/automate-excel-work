import pandas as pd
import numpy as np

''' Set dataframe options '''

pd.options.mode.chained_assignment = None
pd.set_option('display.max_columns', 30)
pd.set_option('display.max_rows', 55)


''' Read input file '''

xnk_url, xnk_header = 'data/posco/xnk.xls', 9
take_out_url, take_out_header = 'data/posco/take_out.xlsx', 9
take_out_free_url, take_out_free_header = 'data/posco/take_out_free.xls', 6
return__url, return_header = 'data/posco/return.xlsx', 6
# warehouse_url, warehouse_header = 'data/posco/iob.xlsx', 9
transfer_url, transfer_header = 'data/posco/transfer.xlsx', 6
last_year_url, last_year_header = 'data/posco/e31_lastyear.xlsx', 11

xnk = pd.read_excel(xnk_url, header=xnk_header)
take_out = pd.read_excel(take_out_url, header=take_out_header)
take_out_free = pd.read_excel(take_out_free_url, header=take_out_free_header)
return_df = pd.read_excel(return__url, header=return_header)
# warehouse = pd.read_excel(
#     warehouse_url, header=warehouse_header, skipfooter=10)
transfer = pd.read_excel(transfer_url, header=transfer_header)
last_year = pd.read_excel(
    last_year_url, sheet_name='BCQT.15', header=last_year_header
)


''' Take E31, B13, A42 from xnk file '''

e31 = xnk.loc[xnk['Mã loại hình'] == 'E31', [
    'Mã NPL/SP', 'Ngày ĐK', 'Đơn vị tính', 'Tổng số lượng']]
b13 = xnk.loc[xnk['Mã loại hình'] == 'B13', [
    'Mã NPL/SP', 'Ngày ĐK', 'Đơn vị tính', 'Tổng số lượng']]
# a42 = xnk.loc[xnk['Mã loại hình'] == 'A42', [
#     'Mã NPL/SP', 'Ngày ĐK', 'Đơn vị tính', 'Tổng số lượng']]
a42 = pd.read_excel(
    'data/posco/a42.xls', header=9
)
a42 = a42.loc[a42['Ngày ĐK'] == '2020-12-31']
# a42['Mã NPL/SP'] = [x[1] for x in a42['Tên hàng'].str.split('-')]

a42 = a42[['Mã NPL/SP', 'Ngày ĐK', 'Đơn vị tính', 'Tổng số lượng']]


for df in [e31, b13, a42]:
    df.rename(columns={
        'Mã NPL/SP': 'Items',
        'Ngày ĐK': 'Date',
        'Đơn vị tính': 'Unit',
        'Tổng số lượng': 'Qty'
    }, inplace=True)


''' Output for production '''

take_out_free = take_out_free[[
    'Date', 'Item Code', 'Item Name', 'Unit', 'Qty', 'Weight']]
take_out_free['Date'] = pd.to_datetime(
    take_out_free['Date'], format="%d/%m/%Y", yearfirst=True
)

del take_out['Date']

take_out.rename(columns={
    'Date.1': 'Date'
}, inplace=True)
take_out['Date'] = pd.to_datetime(
    take_out['Date'], format="%d/%m/%Y", yearfirst=True)
take_out = take_out[['Date', 'Item Code',
                     'Item Name', 'Unit', 'Qty', 'Weight']]

output_for_production = pd.concat([take_out, take_out_free])
output_for_production.reset_index(inplace=True)
del output_for_production['index'], take_out, take_out_free
output_for_production.rename(columns={
    'Item Code': 'Items'
}, inplace=True)


''' Return '''

return_df = return_df[['Date', 'Item Code',
                       'Unit', 'Return Qty', 'Return Weight']]
return_df['Date'] = pd.to_datetime(
    return_df['Date'], format="%Y%m%d", yearfirst=True)

return_df.rename(columns={
    'Item Code': 'Items',
    'Return Qty': 'Qty',
    'Return Weight': 'Weight'
}, inplace=True)

return_df.fillna(0, inplace=True)
return_df['Weight'].replace('-', 0, inplace=True)
return_df['Weight'].replace(' ', 0, inplace=True)
return_df['Weight'] = [float(x) for x in return_df['Weight']]


''' WAREHOUSE '''

# warehouse.rename(columns={
#     'Unnamed: 0': 'Items',
#     'Unnamed: 1': 'Name',
#     'Unnamed: 3': 'Unit',
#     'Unnamed: 2': 'Project',
#     'Unnamed: 13': 'Qty'
# }, inplace=True)
# warehouse = warehouse[['Items', 'Project', 'Unit', 'Qty']]


''' TRANSFER '''

transfer['Date'] = pd.to_datetime(
    transfer['Date'], format="%d/%m/%Y", yearfirst=True)
transfer = transfer[['Date', 'Item Code',
                     'Unit', 'Transfer Qty', 'Transfer Weight']]
transfer.rename(columns={
    'Item Code': 'Items',
    'Transfer Qty': 'Qty',
    'Transfer Weight': 'Weight'
}, inplace=True)


''' GENERATE RESULT '''

# Prepare base items to fill data
base_report = e31.drop_duplicates(subset=['Items'])
base_report = base_report[['Items', 'Unit']]
base_report = pd.merge(
    base_report, output_for_production[['Items', 'Item Name']].drop_duplicates(subset=['Items']).rename(
        {'Unit': 'Output Unit'}
    ),
    how='left', left_on='Items', right_on='Items'
)
base_report = base_report.drop_duplicates(subset=['Items'])

last_year.rename(columns={
    'Unnamed: 1': 'Items',
    'Unnamed: 2': 'Item Name',
    ' - Ngày: 19/02/2020 - Ngày hết hạn: 31/12/2020': 'Unit',
    'Unnamed: 10': 'Ecus stock'
}, inplace=True)
e31_last_year = last_year[['Items', 'Item Name', 'Unit']]

base_report = base_report[['Items', 'Item Name', 'Unit']]
base_report = pd.merge(
    base_report, e31_last_year, how='outer',
    left_on=['Items'],
    right_on=['Items']
)
base_report['Item Name'] = np.where(
    (
        (base_report['Item Name_x'].notnull())
    ),
    base_report['Item Name_x'],
    base_report['Item Name_y']
)
base_report['Unit'] = np.where(
    (
        (base_report['Unit_x'].notnull())
    ),
    base_report['Unit_x'],
    base_report['Unit_y']
)
base_report.drop(['Unit_x', 'Unit_y', 'Item Name_x',
                 'Item Name_y'], axis=1, inplace=True)

last_year = last_year[['Items', 'Ecus stock']]


# List of dataframes for 12 month 1 --> 12
E31_DF = [None] * 13
B13_DF = [None] * 13
A42_DF = [None] * 13
OUTPUT_DF = [None] * 13
RETURN_DF = [None] * 13
TRANSFER_DF = [None] * 13
WAREHOUSE_DF = [None] * 13
BALANCED_REPORT = [None] * 13


for month in range(1, 13):

    E31_DF[month] = e31.loc[e31['Date'].dt.month == month]
    B13_DF[month] = b13.loc[b13['Date'].dt.month == month]
    if month == 1:
        A42_DF[month] = a42
    else:
        A42_DF[month] = a42.loc[((b13['Date'].dt.year == 2021) & (
            a42['Date'].dt.month == month))]
    OUTPUT_DF[month] = output_for_production.loc[output_for_production['Date'].dt.month == month]
    RETURN_DF[month] = return_df.loc[return_df['Date'].dt.month == month]
    TRANSFER_DF[month] = transfer.loc[transfer['Date'].dt.month == month]

    # Prepare for merge dataframe latter.
    E31_DF[month] = E31_DF[month].groupby(
        ['Items'], as_index=False)[['Qty']].sum()
    E31_DF[month].rename(columns={
        'Qty': 'Import',
    }, inplace=True)

    B13_DF[month] = B13_DF[month].groupby(
        ['Items'], as_index=False)[['Qty']].sum()
    B13_DF[month].rename(columns={
        'Qty': 'Re-export',
    }, inplace=True)

    A42_DF[month] = A42_DF[month].groupby(
        ['Items'], as_index=False)[['Qty']].sum()
    A42_DF[month].rename(columns={
        'Qty': 'Re-purpose',
    }, inplace=True)

    OUTPUT_DF[month] = OUTPUT_DF[month].groupby(['Items'], as_index=False)[
        ['Qty', 'Weight']].sum()
    OUTPUT_DF[month].rename(columns={
        'Qty': 'Output for production',
        'Weight': 'Output Weight'
    }, inplace=True)

    RETURN_DF[month]['Qty'] = [int(x) for x in RETURN_DF[month]['Qty']]
    RETURN_DF[month] = RETURN_DF[month].groupby(['Items'], as_index=False)[
        ['Qty', 'Weight']].sum()
    RETURN_DF[month].rename(columns={
        'Qty': 'Return',
        'Weight': 'Return Weight'
    }, inplace=True)

    TRANSFER_DF[month] = TRANSFER_DF[month].groupby(['Items'], as_index=False)[
        ['Qty', 'Weight']].sum()
    TRANSFER_DF[month].rename(columns={
        'Qty': 'Transfer',
        'Weight': 'Transfer Weight'
    }, inplace=True)

    # Merge dataframes
    BALANCED_REPORT[month] = base_report

    if month > 1:
        last_ecus = BALANCED_REPORT[month - 1][['Items', 'Ecus stock']]
        last_ecus.rename(columns={
            'Ecus stock': 'Begin'
        }, inplace=True)
    else:
        last_ecus = last_year
        last_ecus.rename(columns={
            'Ecus stock': 'Begin'
        }, inplace=True)

    BALANCED_REPORT[month] = pd.merge(
        BALANCED_REPORT[month], last_ecus, how='left',
        left_on=['Items'],
        right_on=['Items']
    )

    BALANCED_REPORT[month] = pd.merge(
        BALANCED_REPORT[month], E31_DF[month], how='left',
        left_on=['Items'],
        right_on=['Items']
    )

    BALANCED_REPORT[month] = pd.merge(
        BALANCED_REPORT[month], B13_DF[month], how='left',
        left_on=['Items'],
        right_on=['Items']
    )

    BALANCED_REPORT[month] = pd.merge(
        BALANCED_REPORT[month], A42_DF[month], how='left',
        left_on=['Items'],
        right_on=['Items']
    )
    BALANCED_REPORT[month] = pd.merge(
        BALANCED_REPORT[month], OUTPUT_DF[month], how='left',
        left_on=['Items'],
        right_on=['Items']
    )
    BALANCED_REPORT[month] = pd.merge(
        BALANCED_REPORT[month], RETURN_DF[month], how='left',
        left_on=['Items'],
        right_on=['Items']
    )

    BALANCED_REPORT[month].fillna(0, inplace=True)
    BALANCED_REPORT[month]['Ecus stock'] = (
        BALANCED_REPORT[month]['Begin'] + BALANCED_REPORT[month]['Import'] - BALANCED_REPORT[month]['Output for production'] -
        BALANCED_REPORT[month]['Re-export'] -
        BALANCED_REPORT[month]['Re-purpose']
    )

    # Read from iob excel and take out that closing balance
    try:
        iob = pd.read_excel('data/posco/iob.xlsx', header=9,
                            skipfooter=10, sheet_name=f'IOB {month:02}2021')
        iob = iob.rename(columns={
            'Unnamed: 13': 'Qty',
            'Unnamed: 0': 'Items'
        })[['Items', 'Qty']]
        iob = iob.groupby(['Items'], as_index=False)[['Qty']].sum()
        iob.rename(columns={
            'Qty': 'Warehouse stock'
        }, inplace=True)
        BALANCED_REPORT[month] = pd.merge(
            BALANCED_REPORT[month], iob, how='left',
            left_on=['Items'],
            right_on=['Items']
        )
        BALANCED_REPORT[month]['Balance'] = BALANCED_REPORT[month]['Warehouse stock'] - \
            BALANCED_REPORT[month]['Ecus stock']
    except ValueError:
        print('No Warehouse sheet found in month ' + str(month))

    BALANCED_REPORT[month] = pd.merge(
        BALANCED_REPORT[month], TRANSFER_DF[month], how='left',
        left_on=['Items'],
        right_on=['Items']
    )

    BALANCED_REPORT[month]['Output for production'] = np.where(
        (BALANCED_REPORT[month]['Unit'] == 'Kilogam'), BALANCED_REPORT[month]['Output Weight'], BALANCED_REPORT[month]['Output for production'])

    BALANCED_REPORT[month]['Return'] = np.where(
        (
            (BALANCED_REPORT[month]['Unit'] == 'Kilogam')
        ),
        BALANCED_REPORT[month]['Return Weight'],
        BALANCED_REPORT[month]['Return']
    )

    BALANCED_REPORT[month]['Transfer'] = np.where(
        (
            (BALANCED_REPORT[month]['Unit'] == 'Kilogam')
        ),
        BALANCED_REPORT[month]['Transfer Weight'],
        BALANCED_REPORT[month]['Transfer']
    )

with pd.ExcelWriter('output/posco_result_2.xlsx', engine='xlsxwriter') as writer:
    for df in enumerate(BALANCED_REPORT[1:], start=1):
        df[1].drop(['Return Weight', 'Output Weight',
                   'Transfer Weight'], axis=1, inplace=True)
        df[1].fillna(0, inplace=True)

        df[1].to_excel(writer, sheet_name=f'BALANCED_{(df[0]):02}')

        workbook = writer.book
        worksheet = writer.sheets[f'BALANCED_{(df[0]):02}']

        # Add a format. Light red fill with dark red text.
        format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                       'font_color': '#9C0006',
                                       'num_format': '#,##0.00'})

        # Set the conditional format range.
        start_row = 1
        start_col = 4
        end_row = len(df[1])
        end_cold = 14

        # Apply a conditional format to the cell range.
        worksheet.conditional_format(start_row, start_col, end_row, end_cold,
                                     {'type':     'cell',
                                      'criteria': '<',
                                      'value':    0,
                                      'format':   format1})
print('DONE')
