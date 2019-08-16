#%%
import pandas as pd
import shutil
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from excel_helper import convert_ws_df, refresh_pv

#merge new data
sales_df = pd.read_excel('files/100_Sales_Records.xlsx', 
                         sheet_name='100_Sales_Records', 
                         encoding='ISO-8859-1')
new_sales_df = pd.read_csv('files/1000_Sales_Records.csv', encoding='ISO-8859-1')
merged_df = pd.concat([sales_df, new_sales_df])

#create a copy
shutil.copy2('files/100_Sales_Records.xlsx', 'files/New_Sales_Records.xlsx')

#update worksheet
wb = load_workbook(filename = 'files/New_Sales_Records.xlsx')
sales_ws = wb['100_Sales_Records']
rows = dataframe_to_rows(merged_df, index=False)
for r_idx, row in enumerate(rows, 1):
    for c_idx, value in enumerate(row, 1):
         sales_ws.cell(row=r_idx, column=c_idx, value=value)

#refresh pivot table
refresh_pv(wb['PVT'])

#save updates
wb.save('files/New_Sales_Records.xlsx')

#%%
import pandas as pd
import shutil
import math
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils.cell import get_column_letter
from openpyxl.formula.translate import Translator
from itertools import islice
from excel_helper import convert_ws_df, refresh_pv

#merge new data
wb = load_workbook(filename = 'files/100_Sales_Records_Formula.xlsx')
sales_df = convert_ws_df(wb['100_Sales_Records'], True)
new_sales_df = pd.read_csv('files/1000_Sales_Records.csv', encoding='ISO-8859-1')
merged_df = pd.concat([sales_df, new_sales_df], sort=False)

#create a copy
shutil.copy2('files/100_Sales_Records_Formula.xlsx', 'files/New_Sales_Records_Formula.xlsx')

#update worksheet with formula
wb = load_workbook(filename = 'files/New_Sales_Records_Formula.xlsx')
sales_ws = wb['100_Sales_Records']
rows = dataframe_to_rows(merged_df, index=False)
for r_idx, row in enumerate(rows, 1):
    for c_idx, value in enumerate(row, 1):
        if isinstance(value, float) and math.isnan(value):
            origin_cell_idx = str(get_column_letter(c_idx)) + str(r_idx-1)
            target_cell_idx = str(get_column_letter(c_idx)) + str(r_idx)
            print('r_idx:{},c_idx:{},target_cell_idx:{}'.format(r_idx, c_idx, target_cell_idx))
            value = Translator(sales_ws.cell(row=r_idx-1,column=c_idx).value, origin=origin_cell_idx).translate_formula((target_cell_idx))
        sales_ws.cell(row=r_idx, column=c_idx, value=value)

#refresh pivot table
refresh_pv(wb['PVT'])

#save updates
wb.save('files/New_Sales_Records_Formula.xlsx')

#%%
import pandas as pd
import shutil
import math
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils.cell import get_column_letter
from openpyxl.formula.translate import Translator
from itertools import islice
from excel_helper import convert_ws_df, refresh_pv

#merge new data
wb = load_workbook(filename = 'files/100_Sales_Records_Formula_Cross.xlsx')
sales_df = convert_ws_df(wb['100_Sales_Records'], True)
new_sales_df = pd.read_csv('files/1000_Sales_Records.csv', encoding='ISO-8859-1')
merged_df = pd.concat([sales_df, new_sales_df], sort=False)

#create a copy
shutil.copy2('files/100_Sales_Records_Formula_Cross.xlsx', 'files/New_Sales_Records_Formula_Cross.xlsx')

#update worksheet without formula
wb = load_workbook(filename = 'files/New_Sales_Records_Formula_Cross.xlsx')
sales_ws = wb['100_Sales_Records']
rows = dataframe_to_rows(merged_df, index=False)
for r_idx, row in enumerate(rows, 1):
    for c_idx, value in enumerate(row, 1):
       sales_ws.cell(row=r_idx, column=c_idx, value=value)

#update another worksheet with formula
summary_ws = wb['summary']
c_idx = 1
start_idx = summary_ws.max_row+1
for r_idx in range(start_idx, start_idx+new_sales_df.shape[0]+1):
    origin_cell_idx = str(get_column_letter(c_idx)) + str(r_idx-1)
    target_cell_idx = str(get_column_letter(c_idx)) + str(r_idx)   
    value = Translator(summary_ws.cell(row=r_idx-1,column=c_idx).value, 
                                       origin=origin_cell_idx).translate_formula((target_cell_idx))
    summary_ws.cell(row=r_idx, column=c_idx, value=value)

#refresh pivot table
refresh_pv(wb['PVT'])

#save updates
wb.save('files/New_Sales_Records_Formula_Cross.xlsx')


