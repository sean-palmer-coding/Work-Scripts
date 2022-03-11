import pyodbc as db
import pandas as pd
import os
from datetime import datetime
import xlsxwriter

path = 'C:\\Users\\SPalmer\\OneDrive - CHAS Health\\Desktop\\Temp' #Change path here! use double backslashes.


plist = ['FirstName', 'LastName']                       #variables used to iterate through later
df_list = ['df1', 'df2', 'df3']
df_list1 = []#list instanciated to use later
print("Welcome to Sean's wonderful Browns Consulting Audit Data Program!")
print("Please put the file at this path: " + path)
print("Be sure to check for misspellings in the names. Also, hyphenated names need to be seperated by a space instead.")
datapath = input("Please type in the file name WITHOUT extension (must by xlsx): ")
df = pd.read_excel(os.path.join(path, datapath + ".xlsx"), sheet_name=0, header=4, usecols='B,D',
                   names=plist)               #read providers info out of excel sheet provided by brown consulting
employee_query = 'SELECT FirstName, LastName, EmployeeID, Status FROM Dim_Employee'     #query DB for providers to merge with list provided
conn_string = 'Driver=SQL Server;' \
              'Server=tp-bisql-02;' \
              'Database=CDW;' \
              'Trusted_Connection=yes;'                             #DB connection string
cnxn = db.connect(conn_string)                                      #DB connection object
employees = pd.read_sql(employee_query, cnxn)
df = df.dropna(0, how='all')#query DB for employee data
df = df.apply(lambda x: x.str.strip() if x.dtype == "object" else x)#remove leading and trailing spaces
values = pd.merge(employees, df, on=plist)
missing = df[~df['LastName'].isin(values['LastName']) & ~df['FirstName'].isin(values['FirstName'])].dropna()
missing = missing.merge(employees, on=plist, how='left')#merge the two dataframes
print(values[(values['Status'] == 'Active') | (values['Status'] == 'Leave of absence')].head(20))#Print the active employees after merge to make sure we go who we needed
print('\n\n' + str(len(values[(values['Status'] == 'Active') | (values['Status'] == 'Leave of absence')])) + " Employees were found during the search out of " + str(len(df)) + " listed.")
if len(missing) == 0:
    pass
else:
    print("\n The following provider(s') information is not available (if ID and status are NaN) or no longer an active employee, \ncheck name spelling, or use an alternate provider: \n\n")
    print(missing)
EID = values[(values['Status'] == 'Active') | (values['Status'] == 'Leave of absence')]['EmployeeID'].tolist() #Get employee numbers for string that can be copied into sql if the writing to excel feature fails
EID = [int(x) for x in EID]
var_string = ''.join('(?),' * len(EID))
EID_str = ['(' + str(x) + ')' for x in EID]
print('\n\n')
print("Copy and paste into sql query if desired:")
print(*EID_str, sep=', ')
format_str = ', '.join(EID_str)
cursor = cnxn.cursor()
sqlframe1 = 'SELECT FE.PatientFirst, FE.PatientID, FE.DOB, FE.PatientGender, FE.DOS AS ServiceDate, fb.ProcedureCode, ' \
            'CASE WHEN LEN(fb.Modifier) > 0 THEN LEFT(fb.Modifier, 2) END AS Modifier1, CASE WHEN LEN(' \
            'fb.Modifier) > 3 THEN SUBSTRING(fb.Modifier, 4, 2) END AS Modifier2, CASE WHEN LEN(fb.Modifier) > 6 THEN ' \
            'SUBSTRING(fb.Modifier, 7, 2) END AS Modifier3, FE.Rendering,' \
            'ICD1.DiagnosisCode AS diag1, ICD2.DiagnosisCode AS diag2, ICD3.DiagnosisCode AS diag3, ICD4.DiagnosisCode ' \
            'AS diag4, ICD5.DiagnosisCode AS diag5, ICD6.DiagnosisCode AS diag6, ICD7.DiagnosisCode AS diag7, ' \
            'ICD8.DiagnosisCode AS diag8, ICD9.DiagnosisCode AS diag9, ICD10.DiagnosisCode AS diag10, ' \
            'ICD11.DiagnosisCode AS diag11 FROM Fact_EncounterEZ AS FE LEFT JOIN dbo.Fact_Billing AS fb ON ' \
            'fb.EncounterID = FE.EncID LEFT JOIN (SELECT T.EncounterID, T.DiagnosisCode FROM(SELECT EncounterID,' \
            'DiagnosisCode, ROW_NUMBER() OVER (PARTITION BY EncounterID ORDER BY Ordering) AS RN FROM Fact_Diagnosis) ' \
            'AS T WHERE T.RN = 1 ) AS ICD1 ON ICD1.EncounterID = FE.EncID LEFT JOIN ( SELECT T.EncounterID, ' \
            'T.DiagnosisCode FROM ( SELECT EncounterID, DiagnosisCode, ROW_NUMBER() OVER (PARTITION BY EncounterID ' \
            'ORDER BY Ordering) AS RN FROM Fact_Diagnosis ) AS T WHERE T.RN = 2 ) AS ICD2 ON ICD2.EncounterID = ' \
            'FE.EncID LEFT JOIN ( SELECT T.EncounterID, T.DiagnosisCode FROM ( SELECT EncounterID, DiagnosisCode, ' \
            'ROW_NUMBER() OVER (PARTITION BY EncounterID ORDER BY Ordering) AS RN FROM Fact_Diagnosis ) AS T WHERE ' \
            'T.RN = 3 ) AS ICD3 ON ICD3.EncounterID = FE.EncID LEFT JOIN ( SELECT T.EncounterID, T.DiagnosisCode FROM ' \
            '( SELECT EncounterID, DiagnosisCode, ROW_NUMBER() OVER (PARTITION BY EncounterID ORDER BY Ordering) AS RN ' \
            'FROM Fact_Diagnosis ) AS T WHERE T.RN = 4 ) AS ICD4 ON ICD4.EncounterID = FE.EncID LEFT JOIN ( SELECT ' \
            'T.EncounterID, T.DiagnosisCode FROM ( SELECT EncounterID, DiagnosisCode, ROW_NUMBER() OVER (PARTITION BY ' \
            'EncounterID ORDER BY Ordering) AS RN FROM Fact_Diagnosis ) AS T WHERE T.RN = 5 ) AS ICD5 ON ' \
            'ICD5.EncounterID = FE.EncID LEFT JOIN ( SELECT T.EncounterID, T.DiagnosisCode FROM ( SELECT EncounterID, ' \
            'DiagnosisCode, ROW_NUMBER() OVER (PARTITION BY EncounterID ORDER BY Ordering) AS RN FROM ' \
            'dbo.Fact_Diagnosis ) AS T WHERE T.RN = 6 ) AS ICD6 ON ICD6.EncounterID = FE.EncID LEFT JOIN ( SELECT ' \
            'T.EncounterID, T.DiagnosisCode FROM ( SELECT EncounterID, DiagnosisCode, ROW_NUMBER() OVER (PARTITION BY ' \
            'EncounterID ORDER BY Ordering) AS RN FROM Fact_Diagnosis ) AS T WHERE T.RN = 7 ) AS ICD7 ON ' \
            'ICD7.EncounterID = FE.EncID LEFT JOIN ( SELECT T.EncounterID, T.DiagnosisCode FROM ( SELECT EncounterID, ' \
            'DiagnosisCode, ROW_NUMBER() OVER (PARTITION BY EncounterID ORDER BY Ordering) AS RN FROM Fact_Diagnosis ) ' \
            'AS T WHERE T.RN = 8 ) AS ICD8 ON ICD8.EncounterID = FE.EncID LEFT JOIN ( SELECT T.EncounterID, ' \
            'T.DiagnosisCode FROM ( SELECT EncounterID, DiagnosisCode, ROW_NUMBER() OVER (PARTITION BY EncounterID ' \
            'ORDER BY Ordering) AS RN FROM Fact_Diagnosis ) AS T WHERE T.RN = 9 ) AS ICD9 ON ICD9.EncounterID = ' \
            'FE.EncID LEFT JOIN ( SELECT T.EncounterID, T.DiagnosisCode FROM ( SELECT EncounterID, DiagnosisCode, ' \
            'ROW_NUMBER() OVER (PARTITION BY EncounterID ORDER BY Ordering) AS RN FROM Fact_Diagnosis ) AS T WHERE ' \
            'T.RN = 10 ) AS ICD10 ON ICD10.EncounterID = FE.EncID LEFT JOIN ( SELECT T.EncounterID, T.DiagnosisCode ' \
            'FROM ( SELECT EncounterID, DiagnosisCode, ROW_NUMBER() OVER (PARTITION BY EncounterID ORDER BY Ordering) ' \
            'AS RN FROM Fact_Diagnosis ) AS T WHERE T.RN = 11 ) AS ICD11 ON ICD11.EncounterID = FE.EncID WHERE ' \
            'FE.EmployeeID IN {} AND FE.DOS > DATEADD(DAY, -90, GETDATE()) AND fb.Units_ IS NOT NULL AND fb.Units_ != ' \
            '0;'.format(str((tuple(EID))))
sqlframe2 = 'SELECT Prov.SchedulingName, ProcedureCode, ProcMap.[Procedure Code Description], SUM(Units_) AS Units ' \
            'FROM Fact_Billing Bill LEFT JOIN Dim_Provider Prov ON Bill.ProviderID = Prov.ProviderID ' \
            'LEFT JOIN Athena.[dbo].[Prod_procedurecode] ProcMap ON Bill.ProcedureCode = ProcMap.[Procedure Code] ' \
            'WHERE Bill.EmployeeID IN {} AND [Date] >= DATEADD(MONTH, -6, GETDATE()) GROUP BY Prov.SchedulingName, ' \
            'ProcedureCode, ProcMap.[Procedure Code Description] ORDER BY SUM(Units_) desc '.format(str((tuple(EID))))
sqlframe3 = 'SELECT Prov.SchedulingName, DiagnosisCode, DiagnosisCodeDescription, SUM(Units_) AS Units FROM ' \
            'CDW.dbo.Fact_Billing Bill LEFT JOIN CDW.dbo.Dim_Provider Prov ON Bill.ProviderID = Prov.ProviderID   ' \
            'LEFT JOIN CDW.dbo.Fact_Diagnosis Diag ON Bill.EncounterID = Diag.EncounterID WHERE Bill.EmployeeID IN {} ' \
            "AND Bill.[Date] >= DATEADD(MONTH, -6, GETDATE()) AND ICD9Flag = 'No' GROUP BY Prov.SchedulingName, " \
            'DiagnosisCode, DiagnosisCodeDescription ORDER BY SUM(Units_) desc'.format(str((tuple(EID))))
frame3 = pd.read_sql(sqlframe3, cnxn)   #use above 3 queries to add data into dataframes
frame1 = pd.read_sql(sqlframe1, cnxn)
frame2 = pd.read_sql(sqlframe2, cnxn)
frame1['DOB'] = pd.to_datetime(frame1.DOB).dt.strftime('%m/%d/%Y')                      #format dates
frame1['ServiceDate'] = pd.to_datetime(frame1.ServiceDate).dt.strftime('%m/%d/%Y')
frame1 = frame1.where(pd.notnull(frame1), 'NULL')                                       #replace NaN with NULL string
now = datetime.now()                                                                    # get current datetime for file name
writer = pd.ExcelWriter(os.path.join(path, 'CHAS to BCA - ' + now.strftime('%Y-_%m_%d') + '.xlsx'), engine='xlsxwriter')
frame1.to_excel(writer, sheet_name='Patient Selection', startrow=4, header=False, index=False) #create sheets, write data to them
frame2.to_excel(writer, sheet_name='Dx', index=False)
frame3.to_excel(writer, sheet_name='CPT', index=False)
workbook = writer.book                                                                  #grab workbook object
frame1_worksheet = writer.sheets['Patient Selection']                                   #select active sheet

header_t_format = workbook.add_format({'font_color': 'blue',                            #styles added
                                       'font_name': 'Optima',
                                       'font_size': '14'},
                                      )
header_format = workbook.add_format({'font_color': 'white',
                                     'font_name': 'Calibri'})
header_t_format1 = workbook.add_format({'font_color': 'green',
                                        'font_name': 'Optima',
                                        'font_size': '14'})
dateformat = workbook.add_format()
header_format.set_bg_color('black')
frame1_worksheet.freeze_panes(4, 1)
frame1_worksheet.set_column('C:C')
frame1_worksheet.write(0, 0, "The report should be for the previous 90 days for all "
                             "CPT codes and HCPCS codes for each clinician included in the project.",
                       header_t_format)
frame1_worksheet.write(1, 0, "BCA must receive the report with formatting exactly like the example below.",
                       header_t_format1)
column_list = ['Patient First', 'Patient ID', 'DOB', 'Birth Gender', 'Service Date', 'CPT Code', 'Modifier1',
               'Modifier2', 'Modifier3', 'Doctor']
for x in range(0, 20):
    column_list.append('Diagnosis' + str(x))                                    #creating header for Patient Selection sheet
for col in range(1, (len(column_list)) + 1):
    frame1_worksheet.write(3, col - 1, column_list[col - 1], header_format)
writer.save()
print('Done')                                                                   #save and exit
quit()
