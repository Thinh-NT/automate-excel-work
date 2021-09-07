for name, sheet in goods_processing.items():
        sheet = sheet[['Unnamed: 1', 'Unnamed: 12']]
        sheet.rename(columns={
            'Unnamed: 1': 'Trading', 'Unnamed: 12': 'Processing',
        }, inplace=True)
        if material.value in sheet['Trading'].unique():
            print(
                f'Item {material} in trading of {name}')
            material.trading = True

        elif material in sheet['Processing'].unique():
            print(
                f'Item {material} in processing of {name}')
            material.processing = True
        else:
            pass
    if material.trading and not material.processing:
        print(f'{material.value} is in trading --> move item to RAW and remove from FG')
        raw_concat = raw_concat.append(
            fg_concat[fg_concat['Material'] == material])
        fg_concat = fg_concat.drop(
            fg_concat[fg_concat['Material'] == material].index)
        check_case_one_items.remove(material.value)
    elif material.processing and not material.trading:
        print(
            f'{material.value} is in processing --> move item to FG and remove from RAW')
        fg_concat = fg_concat.append(
            raw_concat[raw_concat['Material'] == material])
        raw_concat = raw_concat.drop(
            raw_concat[raw_concat['Material'] == material].index)
        check_case_one_items.remove(material.value)
    elif material.trading and material.processing:
        print(
            f'******** {material.value} is in trading and processing, need to check ********')
    else:
        print(f'{material.value} can not be found in trading and processing.')
