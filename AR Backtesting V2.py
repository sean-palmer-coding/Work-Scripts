import numpy as np
import pyodbc as db
import pandas as pd
import os
from datetime import datetime

#SQL connection string and query strings
fd = open('SQLQuery1.sql', 'r')
sql_statement1 = fd.read()
fd.close()
fd = open('SQLQuery2.sql', 'r')
sql_statement2 = fd.read()
fd.close()
conn_string = 'Driver=SQL Server;' \
              'Server=tp-bisql-02;' \
              'Database=Finance;' \
              'Trusted_Connection=yes;'
cnxn = db.connect(conn_string)

#End sql strings


def main():
    """Main function that begins the process of backtesting"""
    balance_table = pd.read_sql(sql_statement2, cnxn)
    balance_table = balance_table.set_index(
        ['ARDate', 'EncounterFinancialClass', 'LineFinancialClass', 'ARDaysBucket']
    )

    rates_table = pd.read_sql(sql_statement1, cnxn)
    rates_table['ARDate'] = pd.to_datetime(rates_table['ARDate'])
    rates_table = rates_table.set_index(
        ['ARDate', 'EncounterFinancialClass', 'LineFinancialClass', 'ARDaysBucket']
    )
    rates_table['ExpectedLossRate'] = -1 - (
            rates_table['Numerator'] / rates_table['Denominator']
    )
    rates_table['ExpectedLossRate'] = rates_table['ExpectedLossRate'].replace([np.inf, -np.inf], -1)
    rates_table = rates_table.join(balance_table, on=['ARDate', 'EncounterFinancialClass', 'LineFinancialClass', 'ARDaysBucket'])
    rates_table_output = rates_table.reset_index().set_index(['ARDate', 'EncounterFinancialClass', 'LineFinancialClass', 'ARDaysBucket', 'Lookback', 'TimePeriod'])
    rates_table_output = rates_table_output['ExpectedLossRate'].fillna(-1).unstack(level='ARDate')
    rates_table = rates_table[rates_table['Lookback'].notna()]
    rates_table['ExpectedLossRate'] = rates_table['ExpectedLossRate'].fillna(-1)
    rates_table['LineAmount'] = rates_table['LineAmount'].fillna(0)
    rates_table['MaxLoss$'] = rates_table['MaxLoss$'].fillna(0)
    rates_table['Allowance'] = rates_table['LineAmount'] * rates_table['ExpectedLossRate']
    rates_table['Error'] = rates_table['Allowance'] - rates_table['MaxLoss$']
    rates_table.reset_index().to_excel('All Date1.xlsx')
    allowance_table = rates_table.reset_index().groupby(
        ['LineFinancialClass', 'ARDate', 'Lookback', 'TimePeriod']
    ).sum()
    maxloss_table_final = balance_table[['MaxLoss$', 'LineAmount']].groupby(['ARDate']).sum().T
    allowance_table_final = allowance_table['Allowance'].unstack(level='ARDate')
    allowance_table_final.to_excel('allowance disag4.xlsx')
    error_table = allowance_table['Error'].unstack(level='ARDate')
    error_table.to_excel('error disag4.xlsx')
    maxloss_table_final.to_excel('balances.xlsx')


main()
