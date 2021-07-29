Input files:
![Input-file](https://res.cloudinary.com/thinhnt/image/upload/v1627440258/uni/inputs_r6kh9w.png)

# POSCO:

- South Korean steel-making company headquartered in Pohang, South Korea

## STEPS IN BALANCE.E21:

- Take items in xnk.xlsx and divide it to three table base on type: E21, B13, A42.
- Merge rows with same ITEM CODE and in column 'Tổng số lượng' using sum as aggregation functions (
  E21: 'Import column',
  B13: 'Re-Export',
  A42: 'Re-Purpose',
  )
- 'Begin' column: equal to Ecus Stock last month.
- 'Output for Production', 'Warehouse' and 'Return' columns from sum of 'Weight' column in 'takeout.xlsx', 'iob.xlsx' and 'return.xlsx' ('Warehouse' taken from 'Closing Balance') (If 'Unit' is 'KG' take 'Weight' else take 'Qty')
- Change value in 'Unit' column base on these conditions:
  - 'Bottle', 'Box', 'CAN', 'EA', 'PCS', 'PAIL' --> 'Cái/Chiếc'
  - 'KG' --> 'Kilogam'
  - 'SH', 'ROLL', 'COIL', 'L', 'M', 'M2' --> 'Cuộn'
  - 'SET', 'SUITE', 'PAIR' --> 'Bộ'
- Lookup base on month (1 --> 12)
- Merge rows with same 'Items', 'Unit' and take sum of 'Qty' column
- Join tables with those columns name :
  - 'Baocaochitiet' --> 'Import'
  - 'B13' --> 're-export'
  - 'A42' --> 're-purpose'
  - 'take_out', 'take_out_free' --> 'Output for production'
  - 'Return Goods' --> 'Return'
  - 'Transfer' --> 'Transfer'
  - ('Import' - 'Output for production' -
    'Re-export' - 'Re-purpose') --> 'Ecus stock'

## FINAL GOALS:

- Result by Monthly, Quarterly, Halfyear, Year with type E21, E31: balance sheet
- Find reasons if Ecus Balance is negative.
- Discrepancy between tables
