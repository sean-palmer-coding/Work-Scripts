import pandas as pd
import os

path = 'C:\\Users\\SPalmer\\OneDrive - CHAS Health\\Desktop\\Temp\\AR Allowance backtesting.xlsx'
path2 = 'C:\\Users\\SPalmer\\OneDrive - CHAS Health\\Desktop\\Temp\\'


actuals_bd_rate = pd.read_excel(path, sheet_name='Sheet3 (2)', header=3).set_index(['EncounterFinancialClass', 'ARDaysBucket'])
actuals_rev_rate = pd.read_excel(path, sheet_name='Sheet3', header=3).set_index(['EncounterFinancialClass', 'ARDaysBucket'])
actuals_dollars_bd = pd.read_excel(path, sheet_name='Sheet2', header=3).set_index(['EncounterFinancialClass', 'ARDaysBucket'])
actuals_dollars_rev = pd.read_excel(path, sheet_name='Sheet4', header=3).set_index(['EncounterFinancialClass', 'ARDaysBucket'])
max_loss_bd = pd.read_excel(path, sheet_name='Sheet5 (2)', header=3).fillna(0).set_index(['EncounterFinancialClass', 'ARDaysBucket'])
max_loss_rev = pd.read_excel(path, sheet_name='Sheet5', header=3).fillna(0).set_index(['EncounterFinancialClass', 'ARDaysBucket'])
actuals_bd_rate = actuals_bd_rate.drop(columns=['Grand Total'])
actuals_rev_rate = actuals_rev_rate.drop(columns=['Grand Total'])
ARdates = actuals_bd_rate.columns

rates = pd.DataFrame()
ARdates_list = list(ARdates)[::-1]

rates_dict = {}


def generate_range(dates_back_start, dates_averaged_end):
    for date in ARdates_list:
        if dates_averaged_end < (len(ARdates_list) - ARdates_list.index(date)):
            idx = len(ARdates_list) - ARdates_list.index(date)
            mean_bd = actuals_bd_rate[ARdates[idx - dates_averaged_end:idx - dates_back_start]]
            mean_bd = mean_bd.mean(
                axis=1).reset_index(name=date).set_index(['EncounterFinancialClass', 'ARDaysBucket'])
            mean_rev = actuals_rev_rate[ARdates[idx - dates_averaged_end:idx - dates_back_start]]
            mean_rev = mean_rev.mean(
                axis=1).reset_index(name=date).set_index(['EncounterFinancialClass', 'ARDaysBucket'])
            if ARdates_list.index(date) == 0:
                rates_bd = mean_bd
                rates_rev = mean_rev
            else:
                rates_bd = pd.concat([rates_bd, mean_bd], axis=1)
                rates_rev = pd.concat([rates_rev, mean_rev], axis=1)
    for column in rates_bd.columns:
        rates_bd[column] = rates_bd[column].fillna(-1)
    for column in rates_rev.columns:
        rates_rev[column] = rates_rev[column].fillna(-1)
    return {
        "Bad Debt Rates": rates_bd,
        "Revenue Rates": rates_rev
}


for i in [14, 20]: #range(1, 6):
    rates_dict['lookback|' + (str(i))] = {}
    for x in range(1, 24):
        if x > i:
            rates_dict['lookback|' + (str(i))]['num_months|' + str(x - i)] = generate_range(i, x)
        else:
            continue

estimates_dict = {}
estimates_dict_bypayer = {}

for i in rates_dict.keys():
    estimates_dict[i] = {}
    estimates_dict_bypayer[i] = {}
    for x in rates_dict[i].keys():
        if i == 'lookback|2':
            if x == 'num_months|3':
                print('Here')
        actuals_dollars_bd_temp = actuals_dollars_bd[rates_dict[i][x]["Bad Debt Rates"].columns].fillna(0)
        actuals_dollars_rev_temp = actuals_dollars_rev[rates_dict[i][x]["Revenue Rates"].columns].fillna(0)
        rates1 = rates_dict[i][x]["Bad Debt Rates"].fillna(-1)
        rates2 = rates_dict[i][x]["Revenue Rates"].fillna(-1)
        max_loss_bd_temp = max_loss_bd[rates_dict[i][x]["Bad Debt Rates"].columns].fillna(0)
        max_loss_rev_temp = max_loss_rev[rates_dict[i][x]["Revenue Rates"].columns].fillna(0)
        result_dollars_bd_temp = actuals_dollars_bd_temp.mul(rates_dict[i][x]["Bad Debt Rates"]).drop('Grand Total', level='EncounterFinancialClass')
        result_dollars_rev_temp = actuals_dollars_rev_temp.mul(rates_dict[i][x]["Revenue Rates"]).drop('Grand Total', level='EncounterFinancialClass')
        estimates_dict_bypayer[i][x] = {'Bad Debt Rates': rates_dict[i][x]["Bad Debt Rates"],
                                        "Revenue Rates": rates_dict[i][x]["Revenue Rates"],
                                        'Revenue': result_dollars_rev_temp,
                                        'Bad Debt': result_dollars_bd_temp,
                                        'Total': result_dollars_rev_temp.add(
                                            result_dollars_bd_temp),
                                        'Revenue Δ': result_dollars_rev_temp.subtract(
                                            max_loss_rev_temp
                                        ),
                                        'Bad Debt Δ': result_dollars_bd_temp.subtract(
                                            max_loss_bd_temp
                                        ),
                                        'Total Δ': result_dollars_bd_temp.subtract(
                                            max_loss_bd_temp
                                        ).add(result_dollars_rev_temp.subtract(
                                            max_loss_rev_temp))
                                        }
        data = {}
        for column in result_dollars_rev_temp.columns:
            sum1 = result_dollars_rev_temp[column].sum()
            sum2 = result_dollars_bd_temp[column].sum()
            t_allowance = sum1 + sum2
            data['Lookback'] = i
            data['Averaged_months'] = x
            data[column] = [t_allowance, ]
        estimates_dict[i][x] = data
counter = 0
for keys in estimates_dict.keys():
    for ikeys in estimates_dict[keys].keys():
        if counter == 0:
            allowance_totals = pd.DataFrame(estimates_dict[keys][ikeys])
            counter += 1
        else:
            allowance_totals = allowance_totals.append(pd.DataFrame(estimates_dict[keys][ikeys]))
allowance_totals = allowance_totals.set_index(['Lookback', 'Averaged_months'])
allowance_estimates = allowance_totals.copy()
max_loss = pd.read_excel(path, sheet_name='Sheet4 (2)', header=4).drop(columns=['Grand Total'])
for column in allowance_totals.columns:
    allowance_totals[column] = allowance_totals[column] - max_loss.iloc[0][column]
summarystats = pd.DataFrame({
    ' Total |Δ| ': allowance_totals.abs().sum(axis=1),
    ' Median |Δ| ': allowance_totals.abs().median(axis=1),
    ' Mean |Δ| ': allowance_totals.abs().mean(axis=1)
})
allowance_totals = pd.concat([allowance_totals, summarystats], axis=1)
writer = pd.ExcelWriter(os.path.join(path2, 'AR-Backtest 11-21.xlsx'), engine='xlsxwriter')
workbook = writer.book
allowance_totals.to_excel(writer, sheet_name='AR-Backtest Delta')
allowance_estimates.to_excel(writer, sheet_name='AR-Backtest Actuals')
for i in estimates_dict_bypayer.keys():
    idx = 0
    for x in estimates_dict_bypayer[i].keys():
        for q in estimates_dict_bypayer[i][x].keys():
            estimates_dict_bypayer[i][x][q].to_excel(writer, sheet_name=i, startrow=idx + 1)
            writer.sheets[i].write_string(idx, 0, q)
            writer.sheets[i].write_string(idx, 1, i + ' ' + x)
            idx += len(estimates_dict_bypayer[i][x][q].index) + 5

writer.save()


