import numpy as np
import pandas as pd
from pandas.io.formats.style import Styler
import os
import pathlib

path = pathlib.Path(__file__).parent.resolve()
flagged_providers = []


class Provider:                                             #provider class to store data on providers

    def __init__(self, name, item, jfam, fte, results, salary):
        self.name = name
        self.data = item
        self.jfam = jfam
        self.results = results
        self.fte = fte
        self.salary = salary


class Instructions:                                         #instructions class to store instructions from companion sheet

    def __init__(self, df):
        self.fte_threshold = df['Value'][0]
        self.fte_perc_delta = df['Value'][1]
        self.fte_primary_threshold = df['Value'][2]
        salarydf = pd.read_excel(os.path.join(path, 'Dental FTE analysis.xlsx'), skiprows=3, usecols='G:H', sheet_name=1)
        self.salarymap = {}
        for i in salarydf.columns:
            self.salarymap[i] = salarydf[i][0] * 1.23


def flag(name, instructions, locations, df):
    fte_threshold = instructions.fte_threshold    #the minimum FTE required to be flagged
    fte_perc_threshold = instructions.fte_primary_threshold #the amount of FTE percentage required to determine if a location is the primary location of the provider (to prevent double counting)
    dif_threshold = instructions.fte_perc_delta #the difference between UDS% and FTE% required to flag location and provider
    obj_return = {}                             #instantiate dict to return
    item = df.loc[(df.index.get_level_values('Name') == name)]    #grab data to store in provider object

    jfam = item.index.get_level_values('Job Family')[0]
    salary = instructions.salarymap[jfam]
    fte = item.loc[(item.index.get_level_values('Values') == 'Actual FTE2')].iloc[0]['Unnamed: ' + str(len(df.columns) + 2)]
    if fte < fte_threshold:
        return None
    for loc in locations:
        UDS_enc_perc = item.loc[(item.index.get_level_values('Values') == 'UDS Encounters')].iloc[0][loc] #do comparisons to see if location qualifies for flagging
        Actual_FTE_perc = item.loc[(item.index.get_level_values('Values') == 'Actual FTE')].iloc[0][loc]
        UDS2 = item.loc[(item.index.get_level_values('Values') == 'UDS Encounters2')].iloc[0][loc]
        FTE2 = item.loc[(item.index.get_level_values('Values') == 'Actual FTE2')].iloc[0][loc]
        if abs(UDS_enc_perc - Actual_FTE_perc) > dif_threshold:
            saldelta = (salary / 12 * UDS_enc_perc) - (salary / 12 * Actual_FTE_perc)
            item.loc[(jfam, name, 'Salary Δ'), [loc]] = saldelta
        if abs(UDS_enc_perc - Actual_FTE_perc) > dif_threshold and Actual_FTE_perc < fte_perc_threshold:
            if np.isnan(FTE2): #If the FTE amount is 0 and the locaiton is flagged, then the percentage of UDS encounters is multiplied by the total FTE of the provider in order to calculate FTE should have been classed to that location for the provider
                FTE2 = UDS_enc_perc * item.loc[(item.index.get_level_values('Values') == 'Actual FTE2')].iloc[0]['Unnamed: ' + str(len(df.columns) + 2)]
            obj_return[loc] = [UDS_enc_perc, Actual_FTE_perc, UDS2, FTE2, saldelta]
    if len(obj_return) > 0:                 #append to the flagged providers list if the test returned a flagged location/ provider pair
        provider = Provider(name, item, jfam, fte, obj_return, salary) #create provider object
        flagged_providers.append(provider)
    return


def highlight_col(s, c='yellow'):
    return [f'background-color: {c}' for z in s]            #highlight columns (dont ask me how this works)


def format_numbers(s):                                      #not used
    return {"number-format: percent;"}


columns = ['Job Family', 'Name']                            #instantiate columns to be front filled
instructions = Instructions(pd.read_excel(os.path.join(path, 'Dental FTE analysis.xlsx'), skiprows=3, usecols='E:F', sheet_name=1)) #read instructions into instructions object
df = pd.read_excel(os.path.join(path, 'Dental FTE analysis.xlsx'), skiprows=23) #read data into dataframe
totals = df.tail(4)                                                     #grab totals off dataframe and remove them
df = df.head(-4)
for col in columns:                                                     #front fill columns
    df[col].fillna(method='ffill', axis=0, inplace=True)
providers = df['Name'].unique()                                         #get list of providers
locations = list(df.columns[3:])
locations.pop()                                                         #remove total column from locations list
df = df.set_index(['Job Family', 'Name', 'Values'])                     #create multidimensional array
for i in providers:                                                     #iterate through all providers to flag
    flag(i, instructions, locations, df)
resulting_df = flagged_providers[0].data                                #grab the data from the first flagged provider, then append all the others
for i in flagged_providers[1:]:
    resulting_df = resulting_df.append(i.data)
columns = list(resulting_df.columns)                                    #new columns
column_let = 'ABCDEFGHIJKLMNO'                                          #column letters for excel
column_map = {}                                                         #column map (not used)
for i, z in zip(columns, column_let[3:]):
    column_map[i] = z
style = Styler(resulting_df)                                            #instantiate styler object
for i in flagged_providers:
    ind = flagged_providers.index(i) * len(resulting_df.loc[(resulting_df.index.get_level_values('Name') == i.name)]) #index for iterating
    for q in i.results.keys():                                                  #q is keys of results which are flagged locations
        style.apply(highlight_col, axis=0, subset=(style.index[ind: ind + 5], q))           #highlight locations
        style.applymap(lambda x: 'number-format:0.00%;', subset=(style.index[ind: ind + 2], locations))             #format numbers correctly
        style.applymap(lambda x: 'number-format:0.00%;', subset=(style.index[ind: ind + 2], resulting_df.columns.to_list().pop()))
        style.applymap(lambda x: 'number-format:"$"#.00;', subset=(style.index[ind + 4]))
writer = pd.ExcelWriter(os.path.join(path, 'Dental FTE Jun-AugFY21.xlsx'), engine='xlsxwriter')
style.to_excel(writer, sheet_name='Flagged Providers')          #beggin writing process
workbook = writer.book                                          #grab book object
worksheet = writer.sheets['Flagged Providers']                  #grab active sheet
for i in column_let:
    worksheet.set_column(str(i) + ':' + str(i), 25)             #set column widths
ftesum = 0                                                      #instantiate variable for the sum of all misclassed FTE
len_of_table = (len(flagged_providers) *
                len(resulting_df.loc[(resulting_df.index.get_level_values('Name')
                                      == flagged_providers[0].name)])) + 1 #length of table in cells
for prov in flagged_providers:                                  #iterate to get sum of fte
    for i in prov.results.values():
        ftesum += i[3]

idx = pd.MultiIndex.from_arrays([
    ['Current', 'Current', 'Revised', 'Revised', 'Variance', 'Variance'],
    ['Operating Income', 'Margin', 'Operating Income', 'Margin', 'Operating Income', 'Margin'] #add absolute variance??
])
df = pd.read_excel(os.path.join(path, 'Dental FTE analysis.xlsx'), skiprows=3, sheet_name=1)   #read the workbook again for income and margin data
df = df.head(-1)                                                                                #remove footer

slist = []                                                                                     #instantiate a list for iterating and combining dataframe, and one to keep track of total variance
vtotal = 0
location_sums = {x: 0 for x in locations}                                                       #instantiate a dictionary to keep track of totals for each location
for prov in flagged_providers:                                                              #loop through locations and add up location totals
    for i, v in prov.results.items():
        location_sums[i] += v[4]
for i in locations:                                                                         #create series that are concatinated later into a df that has all locations
    try:                                                                                    #the try block is because Admin does not have net income in the cube and needs to be skipped
        coi = df.loc[df['Location'] == i, 'Operating Income/(Loss)'].iloc[0]
        cm = df.loc[df['Location'] == i, 'Gross Margin'].iloc[0]
        voi = location_sums[i]
        rm = cm - location_sums[i] / (coi / cm)
        roi = coi - location_sums[i]
        vm = location_sums[i] / (coi / cm)
        slist.append(pd.Series([
            coi, cm, roi, rm, voi, vm
        ], name=i, index=idx
        ))
        vtotal += abs(voi)
    except IndexError:
        continue
dollardf = slist[0]                                                                         #concat all the series into a dataframe
for i in slist[1:]:
    dollardf = pd.concat([dollardf, i], axis=1)
style2 = Styler(dollardf)
idx_list = list(style2.index)
style2.applymap(lambda x: 'number-format:"$"#,###.00;', subset=(style2.index[0], locations[1:]))#format the numbers correctly
for x in style2.index:
    if not idx_list.index(x) % 2:
        style2.applymap(lambda x: 'number-format:"$"#,###.00;', subset=(x, style2.columns))
    else:
        style2.applymap(lambda x: 'number-format:0.00%;', subset=(x, style2.columns))
print('done')
currency_format = workbook.add_format({'num_format': '$#,##0.00'})
style2.to_excel(writer, sheet_name='Flagged Providers', startrow=len_of_table + 4, startcol=2)  #write dataframe to excel
worksheet.write(len_of_table + 2, 4, "Missclassed FTE total : " + str(ftesum))                  #write instruction data, assumptions, and filters to excel sheet
worksheet.write(len_of_table + 12, 4, "Absolute Value of Missclassed FTE Estimated in $ : ")
worksheet.write(len_of_table + 12, 5, vtotal, currency_format)
worksheet.write(len_of_table + 14, 4, "Filter: Minimum of " + str(instructions.fte_threshold) + "FTE in the time period")
worksheet.write(len_of_table + 16, 4, "Salary assumptions: ")
idx = 17
for i, x in instructions.salarymap.items():
    worksheet.write(len_of_table + idx, 4, str(i) + ": $")
    worksheet.write(len_of_table + idx, 5, x / 1.23, currency_format)
    worksheet.write(len_of_table + idx, 6, "* 23% for benefits")
    worksheet.write(len_of_table + idx, 7, x, currency_format)
    idx += 1
writer.save()
quit()
