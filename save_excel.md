# how to save dataframe using xlsxwriter?

In this article we will solve the following questions:

1. how to use xlsxwriter to save excel?
2. how to save multiple dataframes to worksheets using Python?
3. how to use Python to stylize the excel formatting?
4. how to set specific data type such as datetime /interger/string/floating data?

## Question Description:

please imagine that you have following 2 dataframes and you want to save them into one excel file with different sheets
and stylize the excel formatting

```python
# generate two dataframes 
import pandas as pd

df1 = pd.DataFrame(
    {'Date': ['2022/12/1', '2022/12/1', '2022/12/1', '2022/12/1'], 'Int': [116382, 227393, 3274984, 438164],
     'Int_with_seperator': [1845132, 298145, 336278, 443816], 'String': ['Tom', 'Grace', 'Luca', 'Tessa'],
     'Float': [98.45, 65.24, 30, 80.88], 'Percent': [0.8878, 0.9523, 0.4545, 0.9921]})
df2 = pd.DataFrame({'Date': ['2022/11/1', '2022/11/1', '2022/11/1', '2022/11/1'], 'Int': [233211, 24321, 35345, 23223],
                    'Int_with_seperator': [925478, 23484, 123249, 2345675],
                    'String': ['Apple', 'Huawei', 'Xiaomi', 'Oppo'], 'Float': [98.45, 65.24, 30, 80.88],
                    'Percent': [0.4234, 0.9434, 0.6512, 0.6133]})
print(df1)
print(df2)
```

###### df1:
| Date      |     Int |   Int_with_seperator | String   |   Float |   Percent |
|:----------|--------:|---------------------:|:---------|--------:|----------:|
| 2022/12/1 |  116382 |              1845132 | Tom      |   98.45 |    0.8878 |
| 2022/12/1 |  227393 |               298145 | Grace    |   65.24 |    0.9523 |
| 2022/12/1 | 3274984 |               336278 | Luca     |   30    |    0.4545 |
| 2022/12/1 |  438164 |               443816 | Tessa    |   80.88 |    0.9921 |

##### df2:
| Date      |    Int |   Int_with_seperator | String   |   Float |   Percent |
|:----------|-------:|---------------------:|:---------|--------:|----------:|
| 2022/11/1 | 233211 |               925478 | Apple    |   98.45 |    0.4234 |
| 2022/11/1 |  24321 |                23484 | Huawei   |   65.24 |    0.9434 |
| 2022/11/1 |  35345 |               123249 | Xiaomi   |   30    |    0.6512 |
| 2022/11/1 |  23223 |              2345675 | Oppo     |   80.88 |    0.6133 |



## save excel through xlsxwriter
#### 1. import xlsxwriter package and create excel file
use .Workbook to create a new excel file
```python
# import package
import xlsxwriter

# create excel file
workbook = xlsxwriter.Workbook("new_excel.xlsx")
```

#### 2. create two sheets
use .add_worksheet to create sheet and cus we have two dataframes that need to be saved, so we have to create two sheets
```python
worksheet1 = workbook.add_worksheet('df1_sheet')
worksheet2 = workbook.add_worksheet('df2_sheet')
```


#### 3. set header format and save the header
at first, we set header format and use .write_row to save our header.
.write_row has four arguments, they are: worksheet.write_row(row, col, data, cell_format)
```python
header_format = workbook.add_format({
    'valign': 'top',
    'fg_color': '#002060',
    'border': 1,
    'font_color': 'white'})

worksheet1.write_row(0, 0, df1.columns, header_format)
worksheet2.write_row(0, 0, df2.columns, header_format)
```
then the header will look like this:
![header](.save_excel_images/4dfa410f.png)



#### 4. create some format
```python
# set datetime format "m/d/yy"
format_datetime = workbook.add_format({'border': 1})
format_datetime.set_num_format(14) # based on above table, 14 means "m/d/yy"
format_datetime.set_font_size(12) # set font size

# set General format
format_general = workbook.add_format({'border': 1})
format_general.set_num_format(0) # 0 means general
format_general.set_font_size(12)

# set integer format "0"
format_integer = workbook.add_format({'border': 1})
format_integer.set_num_format(1)
format_integer.set_font_size(12)

# set float format "0.00"
format_float = workbook.add_format({'border':1})
format_float.set_num_format(2)
format_float.set_font_size(12)

# set integer format with thousands separators "#,##0"
format_integer_separator = workbook.add_format({'border': 1})
format_integer_separator.set_num_format(3)
format_integer_separator.set_font_size(12)

# set percent format "0.00%"
format_percent = workbook.add_format({'border':1})
format_percent.set_num_format(10)
format_percent.set_font_size(12)
```

#### 5. save the data into excel with cell_format
the zero index row already saved our header, so we start from one index row(set the first argument to 1)
.write_column(row, column, data, cell_format)
```python
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
```

#### 6. set column width
for better display the data, we can also use .set_column to set the column width.
```python
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
```


#### 7. finish
```python
workbook.close()
```

