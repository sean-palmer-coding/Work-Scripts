import pandas as pd
import os

path = 'C:\\Users\\SPalmer\\OneDrive - CHAS Health\\Desktop\\Temp\\Dental'

df = pd.read_excel(os.path.join(path, 'upload-Nextgen AR Snapshots.xlsx'), sheet_name='Upload 2', header=0)
df['ARDate'] = pd.to_datetime(df['ARDate'])
df.to_excel('output_fixed.xlsx', index=None)