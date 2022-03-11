import pandas as pd
import os
path = "H:\\Finance\\PROCUREMENT\\Shared\\Aurum"
files = os.listdir(path)

dfs = []

for file in files:
    df = pd.read_excel(os.path.join(path, file), skiprows=5)
    df = df[df['Case #'].notna()].loc[:, ~df.columns.str.contains('^Unnamed')]
    dfs.append(df)

df = pd.concat(dfs)
df.to_excel('test1.xlsx', index=False)
