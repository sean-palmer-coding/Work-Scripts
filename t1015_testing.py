import pandas as pd
import pyodbc as db
from os import path as p

path = 'H:\\Finance\\REPORTING & PLANNING\\General\\Compliance Reporting\\UDS\\2021\\9D\\Dental\\'
conn_string = 'Driver=SQL Server;' \
              'Server=tp-bisql-02;' \
              'Database=Finance;' \
              'Trusted_Connection=yes;'
t1015_q = open('t1015.sql', 'r').read()
cnxn = db.connect(conn_string)
cursor = cnxn.cursor()
billing_data = pd.read_sql(t1015_q, cnxn)
nextgen_data = pd.read_excel(p.join(path, 'T1015 1-1-21 - 12-31-21.xlsx'), skiprows=4)
columns_keep = ['Sv It', 'Fin Class', 'Loc Name', 'E/I/A/B', 'Name', 'Proc Dt', 'Dt of Svc', 'Pay Amt', 'Prim Payer']
nextgen_data = nextgen_data[columns_keep]
nextgen_data = nextgen_data.rename(columns={'E/I/A/B': 'EncounterID'})
nextgen_data['EncounterID'] = nextgen_data['EncounterID'] + 990000000
nextgen_data = nextgen_data[nextgen_data['Proc Dt'] > '2/1/2021']
nextgen_jdata = nextgen_data.groupby('EncounterID').sum()
billing_data['EncounterID'] = billing_data['EncounterID']
data = billing_data.merge(nextgen_jdata, on='EncounterID', how='left')
data = data[data['Pay Amt'] != -data['Charges_']]
# data.to_excel(p.join(path, 'non-matching4.xlsx'))



