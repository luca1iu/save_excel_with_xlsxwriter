import pandas as pd

df1 = pd.DataFrame(
    {'Date': ['2022/12/1', '2022/12/1', '2022/12/1', '2022/12/1'], 'Int': [116382, 227393, 3274984, 438164],
     'Int_with_seperator': [1845132, 298145, 336278, 443816], 'String': ['Tom', 'Grace', 'Luca', 'Tessa'],
     'Float': [98.45, 65.24, 30, 80.88], 'Percent': [0.8878, 0.9523, 0.4545, 0.9921]})

df2 = pd.DataFrame({'Date': ['2022/11/1', '2022/11/1', '2022/11/1', '2022/11/1'], 'Int': [233211, 24321, 35345, 23223],
                    'Int_with_seperator': [925478, 23484, 123249, 2345675],
                    'String': ['Apple', 'Huawei', 'Xiaomi', 'Oppo'], 'Float': [98.45, 65.24, 30, 80.88],
                    'Percent': [0.4234, 0.9434, 0.6512, 0.6133]})

import xlsxwriter

# create excel file
workbook = xlsxwriter.Workbook("new_excel.xlsx")

# create two sheets
worksheet1 = workbook.add_worksheet('df1')
worksheet2 = workbook.add_worksheet('df2')

# set header format
header_format = workbook.add_format({
    'valign': 'top',
    'fg_color': '#002060',
    'border': 1,
    'font_color': 'white'})

# start from A1 cell
# worksheet.write_row(row, col, data, cell_format)
worksheet1.write_row(0, 0, df1.columns, header_format)
worksheet2.write_row(0, 0, df2.columns, header_format)

# create some format

# set datetime format "m/d/yy"
format_datetime = workbook.add_format({'border': 1})
format_datetime.set_num_format(14)  # based on above table, 14 means "m/d/yy"
format_datetime.set_font_size(12)  # set font size

# set General format
format_general = workbook.add_format({'border': 1})
format_general.set_num_format(0)  # 0 means general
format_general.set_font_size(12)

# set integer format "0"
format_integer = workbook.add_format({'border': 1})
format_integer.set_num_format(1)
format_integer.set_font_size(12)

# set float format "0.00"
format_float = workbook.add_format({'border': 1})
format_float.set_num_format(2)
format_float.set_font_size(12)

# set integer format with thousands separators "#,##0"
format_integer_separator = workbook.add_format({'border': 1})
format_integer_separator.set_num_format(3)
format_integer_separator.set_font_size(12)

# set percent format "0.00%"
format_percent = workbook.add_format({'border': 1})
format_percent.set_num_format(10)
format_percent.set_font_size(12)

worksheet1.write_column(1, 0, df1.iloc[:, 0], format_datetime)
worksheet1.write_column(1, 1, df1.iloc[:, 1], format_integer)
worksheet1.write_column(1, 2, df1.iloc[:, 2], format_integer_separator)
worksheet1.write_column(1, 3, df1.iloc[:, 3], format_general)
worksheet1.write_column(1, 4, df1.iloc[:, 4], format_float)
worksheet1.write_column(1, 5, df1.iloc[:, 5], format_percent)

worksheet2.write_column(1, 0, df2.iloc[:, 0], format_datetime)
worksheet2.write_column(1, 1, df2.iloc[:, 1], format_integer)
worksheet2.write_column(1, 2, df2.iloc[:, 2], format_integer_separator)
worksheet2.write_column(1, 3, df2.iloc[:, 3], format_general)
worksheet2.write_column(1, 4, df2.iloc[:, 4], format_float)
worksheet2.write_column(1, 5, df2.iloc[:, 5], format_percent)

# set column width
worksheet1.set_column('A:A', 15)
worksheet1.set_column('B:B', 15)
worksheet1.set_column('C:C', 15)
worksheet1.set_column('D:D', 15)
worksheet1.set_column('E:E', 15)
worksheet1.set_column('F:F', 15)

worksheet2.set_column('A:A', 15)
worksheet2.set_column('B:B', 15)
worksheet2.set_column('C:C', 15)
worksheet2.set_column('D:D', 15)
worksheet2.set_column('E:E', 15)
worksheet2.set_column('F:F', 15)

workbook.close()
print("excel save successfullly")
