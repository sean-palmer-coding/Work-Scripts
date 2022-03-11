import pandas as pd
import os
import datetime
import time

path = 'H:\\Finance\\ACCOUNTING\\General\\Month End Close\\2022 FY\\04-February 2022\\Revenue\\NextGen' #The current working Directory
print('Ensure the ARDate field is updated to the first of the month in ' #instructions
      'which you are doing the work - not the month end month')
print("For example, the August book close where you're doing the "
      "work in the first week of September: ARDate = 9/1/20xx")
print("\nYour current working directory is: ")
print(path + "\n")
now = datetime.datetime.now()                                           #current date


def main():
    date_input = input('AR Date (mm/dd/yyyy format)?: ')
    while True:
        try:
            try:
                date_input = datetime.datetime.strptime(date_input, '%m/%d/%Y')             #compare current date to make sure it makes sense for
            except (ValueError, TypeError) as e:                                            #the user input
                if e.__class__.__name__ == "ValueError":                                    #Try block checks for exceptions
                    print('The date you entered is either not a valid date, or does not match the format. Try again.')
                    main()
                elif e.__class__.__name__ == "TypeError":
                    pass
                else:
                    print("You really broke it now... An unknown error occurred")
            file_name = input('What is the file name excluding extension? (Must be .xlsx): ')
            if date_input.month != now.month or date_input.year != now.year or date_input.day != 1:
                while True:
                    print('Are you sure? Date input is from a different month, '
                          'year, or the day entered is not the 1st.\n'
                          'Remember, if you are doing month end, and your folder says '
                          'OCTOBER 20XX, then the AR date would be 11/1/20XX')
                    print('Your current path: ' + path)
                    response = input('Proceed?(Y/N): ')
                    if response == 'Y':
                        response2 = input('If you are sure, type OVERRIDE (case sensitive): ')
                        if response2 == 'OVERRIDE':
                            break
                        else:
                            continue
                    elif response == 'N':
                        main()
                    else:
                        print('Your response was not understood')
                        pass
            mapping(date_input, file_name)
            break
        except Exception as e:
            print('An exception has occurred: ')
            print(e.__class__.__name__)
            print(e)
            print('This is embarrassing, lets try that again shall we?\n\n')
            continue


def mapping(date_input, file_name):
    print('\n\nMapping... BeeBooBEEP!')
    mapping = {'Per Nbr': 'PatientNumber',              #remap the columns so that the output matches the table it will be uploaded to
               'E/I/A/B': 'EncID',
               'Fin Class': 'FinancialClass',
               'Loc Name': 'LocationName',
               'Payer Name': 'PayerName',
               'Rendering': 'RenderingProv',
               'Ln Itm Amt': 'LineAmount',
               'Dt of Svc': 'ServiceDate',
               'Chg Amt': 'ChargeAmount',
               'Adj Amt': 'AdjustmentAmount',
               'Pay Amt': 'PaymentAmount',
               'Proc Dt': 'ProcessDate',
               }
    kept_columns = ['LocationName', 'FinancialClass', 'PayerName', 'RenderingProv', 'PatientNumber', 'EncID',       #Columns to be kept here
                    'LineAmount', 'ARDate', 'ChargeAmount', 'ServiceDate',
                    'AdjustmentAmount', 'PaymentAmount', 'ProcessDate'
                    ]
    df = pd.read_excel(os.path.join(path, file_name + '.xlsx'), sheet_name='Page 1', header=4)                      #read the sheet into pandas
    df['balance'] = df['Ln Itm Amt'].groupby(df['E/I/A/B']).transform('sum')                                        #group the balances by EncID, then calculate the balance to determine
    df['ARDate'] = date_input                                                                                       #if it is net positive or negative
    print("The credit balance this AR month is: ")
    print(df[df.balance < 0.00]['Ln Itm Amt'].sum())                                                                #Print credits for JE29
    df = df[df.balance > 0.00].rename(columns=mapping)[kept_columns]                                                #get positive balances ONLY
    df.to_excel(os.path.join(path, 'AR Positive Balances Only - ' + now.strftime("%m-%d-%Y") + '.xlsx'), index=None)#write to new sheet with the uploadable file
    print('\nCompleted without error')
    time.sleep(5)                                                                                                   #in the working directory
    exit()                                                                                                          #Exit (only works in cmd line, not python shell


main()
