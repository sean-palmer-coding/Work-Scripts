import pandas as pd
import os

path = 'C:\\Users\\SPalmer\\Downloads'
costs = ['40 - Personnel Costs', '50 - Supplies', '80 - Other Expenses', '60 - Other Costs']
indirect_costs = ['Depreciation', 'Admin Distribution', 'Centralized Services Distribution']


df = pd.read_excel(os.path.join(path, 'Forecasts 2022-01-31_10_28_42_PST.xlsx'), sheet_name='Forecasts', index_col=None)
df1 = df[df['GL Category *'].isin(costs) & ~df['Line Item'].isin(indirect_costs)].groupby(df['Costing Center *']).sum()
df = df[df['GL Category *'].isin(costs)].groupby(df['Costing Center *']).sum()


# df.to_excel('Total Expenses.xlsx')
df1.to_excel('Direct Expenses.xlsx')