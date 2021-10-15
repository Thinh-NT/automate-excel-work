import pandas as pd
import numpy as np
import warnings

pd.options.mode.chained_assignment = None

warnings.simplefilter(action='ignore', category=FutureWarning)

item_master_excel_file = '../data/inzi/Item Master_10_07_2021 0850.xls'
material_history_excel_file = '../data/inzi/Material Movement History_10_09_2021 08 oct.xls'
out_put_file = '../output/08 oct.xlsx'

item_master = pd.read_excel(
    item_master_excel_file, index_col=False, header=1
)
material_history = pd.read_excel(
    material_history_excel_file, index_col=False, header=1
)

# Item master excel file must have 3 columns 'Material', 'Material Type', 'Procurement'
temp_item_master = item_master[['Material', 'Material Type', 'Procurement']]

'''
PREPARE FOR STEP ONE
'''
step_one = pd.merge(material_history, temp_item_master,
                    on='Material', how='left')

check_item_master = step_one['Material Type'].isnull()

if check_item_master.sum() > 1:
    print(f'Có {check_item_master.sum() - 1} mã không tìm thấy trong item_master')
else:
    pass


step_one.rename(columns={
    'Reference': 'Order Category',
    'Unnamed: 13': 'Order Data',
    'Unnamed: 14': 'Master Category',
    'Unnamed: 15': 'Master Data',
    'Unnamed: 16': 'Remark'
}, inplace=True)

step_one = step_one.iloc[1:, :]

# pd.set_option('mode.chained_assignment', None)

step_one.loc[step_one['Quantity'] > 0, 'Input'] = step_one['Quantity']
step_one.loc[step_one['Quantity'] > 0, 'Output'] = 0
step_one.loc[step_one['Quantity'] <= 0, 'Input'] = 0
step_one.loc[step_one['Quantity'] <= 0, 'Output'] = -step_one['Quantity']

cols = step_one.columns.tolist()

cols = cols[:10] + [cols[-2], cols[-1]] + cols[10:-2]

step_one = step_one[cols]

step_one['Account code'] = [int(x)
                            for x in (step_one['Movement Type'].str[:3])]

cols = step_one.columns.tolist()
cols = cols[:7] + [cols[-1]] + cols[7: -1]
step_one = step_one[cols]

steady = step_one.pivot_table(index=['Unnamed: 0', 'Material'],
                              columns='Account code', values=['Quantity'], aggfunc=np.sum, fill_value=0)


steady.reset_index(inplace=True)
del steady['Unnamed: 0']
steady = steady.groupby(steady['Material']).aggregate(
    'sum')


steady['IN'] = steady.where(steady > 0).sum(1)
steady['OUT'] = steady.where(steady < 0).sum(1)

SUMMARY = steady.iloc[:, [-2, -1]]


columns_name = {
    'H': 'Account code',
    'p': 'Order Category',
    'W': 'Material Type',
    'X': 'Procurement'
}

'''
Conditions to insert into RAW, WIM, FG
'''

''' SCM RAW '''
nhap_vao = [None] * 9
nhap_vao[0] = ((step_one['Account code'] == 101) & (
    step_one['Order Category'] == 'Purchase Order'))
nhap_vao[1] = (step_one['Account code'] == 103)
nhap_vao[2] = (step_one['Account code'] == 602)
nhap_vao[3] = (step_one['Account code'] == 610)
nhap_vao[4] = (step_one['Account code'] == 623)
nhap_vao[5] = (step_one['Account code'] == 701)
nhap_vao[6] = (step_one['Account code'] == 720)
nhap_vao[7] = (step_one['Account code'] == 801)
nhap_vao[8] = (step_one['Account code'] == 809)

xuat_ra = [None] * 10
xuat_ra[0] = ((step_one['Account code'] == 102) & (
    step_one['Order Category'] == 'Purchase Order'))
xuat_ra[1] = (step_one['Account code'] == 201)
xuat_ra[2] = (step_one['Account code'] == 261)
xuat_ra[3] = (step_one['Account code'] == 555)
xuat_ra[4] = (step_one['Account code'] == 601)
xuat_ra[5] = (step_one['Account code'] == 609)
xuat_ra[6] = (step_one['Account code'] == 702)
xuat_ra[7] = (step_one['Account code'] == 712)
xuat_ra[8] = (step_one['Account code'] == 721)
xuat_ra[9] = (step_one['Account code'] == 803)

thuyen_chuyen = [None]
thuyen_chuyen[0] = (step_one['Account code'].between(300, 402))


def condition_to_data(arr):
    material_type_order = ['ROH', 'HAWA', 'HALB', 'FERT']
    material_order_index = dict(
        zip(material_type_order, range(len(material_type_order))))
    for i in range(len(arr)):
        arr[i] = step_one.loc[arr[i]]
        arr[i]['Tm_Rank'] = arr[i]['Material Type'].map(
            material_order_index)  # Sort rows base on material_type_order
        arr[i].sort_values(['Tm_Rank', 'Procurement'], inplace=True)
        arr[i].drop('Tm_Rank', 1, inplace=True)


condition_to_data(nhap_vao)
condition_to_data(xuat_ra)
condition_to_data(thuyen_chuyen)

step_two = pd.concat(nhap_vao + xuat_ra + thuyen_chuyen)

condition_to_delete = ~(step_two['Account code'].isin([101, 102])) & (
    ((step_two['Material Type'] == 'ROH') & (step_two['Procurement'] == 'E')) |
    ((step_two['Material Type'] == 'HAWA') & (step_two['Procurement'] == 'E')) |
    ((step_two['Material Type'] == 'HALB') & (step_two['Procurement'].isin(['E', 'X']))) |
    ((step_two['Material Type'] == 'FERT') &
     (step_two['Procurement'].isin(['E', 'X'])))
)
step_two.drop(step_two[condition_to_delete].index, inplace=True)

SCM_RAW = step_two  # Get SCM_RAW
'''------------------------------'''

''' SCM WIP '''
nhap_vao = [None] * 2
nhap_vao[0] = (
    (step_one['Account code'] == 101) &
    (step_one['Order Category'] == 'Production Order') &
    (step_one['Material Type'] == 'HALB') &
    (step_one['Procurement'].isin(['E', 'X']))
)

nhap_vao[1] = (
    (step_one['Account code'].isin([103, 602, 610, 623, 701, 720, 801, 809])) &
    (step_one['Material Type'] == 'HALB') &
    (step_one['Procurement'].isin(['E', 'X']))
)

xuat_ra = [None] * 3
xuat_ra[0] = (
    (step_one['Account code'] == 102) &
    (step_one['Order Category'] == 'Production order') &
    (step_one['Material Type'] == 'HALB') &
    (step_one['Procurement'].isin(['E', 'X']))
)

xuat_ra[1] = (
    (step_one['Account code'].isin([201, 555, 601, 609, 702, 712, 721, 803])) &
    (step_one['Material Type'] == 'HALB') &
    (step_one['Procurement'].isin(['E', 'X']))
)

xuat_ra[2] = (
    (step_one['Account code'] == 261) &
    (
        ((step_one['Material Type'] == 'HALB') &
         (step_one['Procurement'].isin(['E', 'X']))) |
        ((step_one['Material Type'] == 'FERT') &
         (step_one['Procurement'].isin(['E', 'X'])))
    )
)

thuyen_chuyen = [None]
thuyen_chuyen[0] = (
    (step_one['Account code'].between(300, 402)) &
    (step_one['Material Type'] == 'HALB') &
    (step_one['Procurement'].isin(['E', 'X']))
)

condition_to_data(nhap_vao)
condition_to_data(xuat_ra)
condition_to_data(thuyen_chuyen)

SCM_WIP = pd.concat(nhap_vao + xuat_ra + thuyen_chuyen)
'''--------------------------------'''

# SCM_WIP.loc[(
#     (SCM_WIP['Account code'] == 261) &
#     (SCM_WIP['Material Type'] == 'FERT') &
#     (SCM_WIP['Location'].str.startswith('RW'))
# ), 'Location'] = None


''' SCM FG '''
nhap_vao = [None] * 2
nhap_vao[0] = (
    (step_one['Account code'] == 101) &
    (step_one['Order Category'] == 'Production Order') &
    (step_one['Material Type'] == 'FERT') &
    (step_one['Procurement'].isin(['E', 'X']))
)

nhap_vao[1] = (
    (step_one['Account code'].isin([103, 602, 623, 701, 720, 801, 809])) &
    (step_one['Material Type'] == 'FERT') &
    (step_one['Procurement'].isin(['E', 'X']))
)

xuat_ra = [None] * 2
xuat_ra[0] = (
    (step_one['Account code'] == 102) &
    (step_one['Order Category'] == 'Production Order') &
    (step_one['Material Type'] == 'FERT') &
    (step_one['Procurement'].isin(['E', 'X']))
)

xuat_ra[1] = (
    (step_one['Account code'].isin([201, 261, 555, 601, 609, 702, 712, 721, 803])) &
    (step_one['Material Type'] == 'FERT') &
    (step_one['Procurement'].isin(['E', 'X']))
)

thuyen_chuyen = [None]
thuyen_chuyen[0] = (
    (step_one['Account code'].between(300, 402)) &
    (step_one['Material Type'] == 'FERT') &
    (step_one['Procurement'].isin(['E', 'X']))
)

condition_to_data(nhap_vao)
condition_to_data(xuat_ra)
condition_to_data(thuyen_chuyen)

SCM_FG = pd.concat(nhap_vao + xuat_ra + thuyen_chuyen)

# with pd.ExcelWriter('output/before_report.xlsx') as writer:
#     SCM_RAW.to_excel(writer, sheet_name='RAW')
#     SCM_FG.to_excel(writer, sheet_name='FG')
#     SCM_WIP.to_excel(writer, sheet_name='WIP')


def take_sum_of_positive_number(x):
    positive_nums = []
    for val in x:
        if val <= 0:
            continue
        else:
            positive_nums.append(val)
    return np.sum(positive_nums)


def get_result(df):
    '''
    Convert to final report, columns base on Account Code
    '''

    aggregation_functions = {}  # Methods use when group by 'Material'
    final_col = []              # Columns will appear in result
    in_columns = []             # Columns will be taken as sum in IN column
    out_columns = []            # Columns will be taken as sum in OUT column

    columns_order = [101, 102, 'PURCHASE', 103, 321, 323, 327, 342, 343, 344, 346, 401, 602, 610, 623, 701,
                     720, 801, 809, 201, 261, 555, 601, 609, 702, 712, 721, 803]
    columns_inin = [101, 103, 602, 610, 621, 623, 701, 720, 801, 809]
    columns_inout = [102, 201, 261, 555, 601, 609, 702, 712, 721, 803]

    for column in pd.unique(df['Account code']):
        if isinstance(column, np.int64):
            aggregation_functions[column] = 'sum'
            if column in [321, 323, 327, 342, 343, 344, 346, 401]:
                aggregation_functions[column] = take_sum_of_positive_number

    result = df.pivot_table(index=['Unnamed: 0', 'Material'],
                            columns='Account code', values='Quantity', aggfunc=np.sum, fill_value=0)

    result.reset_index(inplace=True)

    del result['Unnamed: 0']

    result = result.groupby(result['Material']).aggregate(
        aggregation_functions)

    for column in pd.unique(step_one['Account code']):
        if column not in result.columns:
            result[column] = 0

    if 102 in result:
        if not (result[102] == 0).all():
            result['PURCHASE'] = result[101] - result[102]

    # REARRANGE COLUMNS ORDER
    for col in columns_order:
        if col in result.columns:
            final_col.append(col)

    result = result[final_col]

    for column in result.columns:
        if column in columns_inin:
            in_columns.append(column)
        if column in columns_inout:
            out_columns.append(column)

    result['IN'] = result[in_columns].sum(axis=1)
    result['OUT'] = result[out_columns].sum(axis=1)

    result.fillna(0, inplace=True)
    result.reset_index(inplace=True)
    result.index.name = None

    # Remove columns with no values
    # result = result.loc[:, (result != 0).any(axis=0)]

    for column in result:
        if isinstance(column, int) or column in ['IN', 'OUT']:
            result[column] = result[column].abs()
    return result


RAW_REPORT = get_result(SCM_RAW)
WIP_REPORT = get_result(SCM_WIP)
FG_REPORT = get_result(SCM_FG)

with pd.ExcelWriter(out_put_file) as writer:
    RAW_REPORT.to_excel(writer, sheet_name='REPORT_RAW')
    WIP_REPORT.to_excel(writer, sheet_name='REPORT_WIP')
    FG_REPORT.to_excel(writer, sheet_name='REPORT_FG')
    SUMMARY.to_excel(writer, sheet_name='SUMMARY')

    for sheet in [[RAW_REPORT, 'REPORT_RAW'], [WIP_REPORT, 'REPORT_WIP'], [FG_REPORT, 'REPORT_FG']]:
        workbook = writer.book
        worksheet = writer.sheets[sheet[1]]

        # Add a format. Light red fill with dark red text.
        format1 = workbook.add_format({'bg_color': '#FFC7CE',
                                       'font_color': '#9C0006'})

        # Set the conditional format range.
        start_row = 1
        start_col = 1
        end_row = len(sheet[0])
        end_cold = 23

        # Apply a conditional format to the cell range.
        worksheet.conditional_format(start_row, start_col, end_row, end_cold,
                                     {'type':     'cell',
                                      'criteria': '<',
                                      'value':    0,
                                      'format':   format1})

print('DONE')
