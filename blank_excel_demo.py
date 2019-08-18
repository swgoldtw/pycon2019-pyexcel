
#%%Demo: One dataframe to output different data granularity reports
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

df = pd.read_csv('files/100_Sales_Records.csv', encoding='ISO-8859-1')
summary_df = df.groupby(['Region']).aggregate(np.sum).reset_index().filter(['Region', 'Units Sold', 'Total Profit'])
region_df = df.groupby(['Region', 'Country', 'Item Type', 'Sales Channel']).aggregate(np.sum).reset_index().filter(
                ['Region', 'Country', 'Item Type', 'Sales Channel', 'Units Sold', 'Total Profit'])
summary_df
wb = Workbook()
ws = wb.active
ws.title = 'Summary'
for r in dataframe_to_rows(summary_df, index=False, header=True):
    ws.append(r)

for name, group in region_df.groupby('Region'):
    ws = wb.create_sheet(title=name[:10])
    for r in dataframe_to_rows(group, index=False, header=True):
        ws.append(r)

wb.save('files/manager_report.xlsx')



#%%
