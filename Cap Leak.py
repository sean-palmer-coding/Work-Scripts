import pandas as pd
import os
import datetime


path = 'H:\\Finance\\MANAGED CARE\\Shared\\BI Upload\\Cap Leak\\Source Files\\All Source Files\\'                   #Where it looks for mapping
path2 = 'H:\\Finance\\MANAGED CARE\\Shared\\BI Upload\\Cap Leak\\Upload Files\\'                                    #where it sends output
print('Please make sure your files are in the following format: \nData files: "[MCO NAME (Capitalized)] Cap Leak '
      '[XX (month)] - [XXXX (year)].xlsx \nMapping file: mapping.xlsx\n THESE ARE CASE SENSITIVE')
print('See mapping file template for instructions on mapping column names')


class CapLeak:

    def __init__(self):
        self.mapping = None
        while True:
            try:
                self.main()
            except Exception as e:
                print('\nOperation NOT COMPLETED. See error below:\n')
                print(e)
                print('\n\n')

    def main(self):
        month = int(input('\nMonth number of data (Example: If January, your entry would be "1"): '))               #Collect month number from user
        if month < 10:
            month = '0' + str(month)
        else:
            month = str(month)#append a leading zero to the month number
        year = str(input('Year: '))                                                                                 # collect year number
        self.mapping, mco_list = self.get_mapping()                                                                 #run the funtion that collects the mapping data from the excel template
        df_list = []                                                                                                #instantiate an empty list to store the dataframes in
        for i in mco_list:
            df_list.append(self.drop_columns(self.mapping[i].keys(), i, month, year))                               #loop through the list in order to find the files and add them as a dataframe, then remove unnecessary columns *see def drop_columns
        if len(df_list) < 1:                                                                                        #check to see if there are no dataframes in memory
            raise Exception('No files found! Check month values and file names')
        for d, i in zip(df_list, mco_list):
            try:                                                                                                    #unpack the list of dataframes and list of MCOs together to rename the columns using the mapping created earlier
                d.rename(columns=self.mapping[i], inplace=True)
            except Exception:
                continue
        final_df = pd.concat(df_list, axis=0, ignore_index=True)                                                    #combine all the dataframes together
        final_df.to_excel(os.path.join(path2, "output - " + datetime.datetime.now().strftime(
            "%Y%m%dT%H%M%S") + ".xlsx"), index=False)                                                               #output to xlsx
        print('\n\nRefactor complete, check for missing columns above.\n')
        input('***Press enter to quit***')
        quit()

    def isNaN(self, num):
        return num != num

    def get_mapping(self):
        mapping = pd.read_excel(os.path.join(path, 'mapping.xlsx'))                                                 #open the excel sheet with mapping in it
        map_dict = {}                                                                                               #instantiate an empty dictionary object to be used later
        for i in mapping.columns[1:]:                                                                               #Skips the first column in the mapping excel sheet
            map_dict[i] = {}                                                                                        #create a nested dictionary for each MCO that contains the mapping data
            for u, m in zip(mapping.to_dict()['Upload_Field'].values(), mapping.to_dict()[i].values()):             #construct the dictionary checking for empty column mappings
                if self.isNaN(m):
                    pass
                else:
                    map_dict[i][m] = u
        return map_dict, list(map_dict.keys())                                                                      #return the mapping and the list of MCOs

    def drop_columns(self, columns_remove, mco, month, year):
        checklist = list(columns_remove)                                                                            #create a list to check the expected
        columns_remove = list(columns_remove)                                                                       #turn columns_remove into list
        try:
            df = pd.read_excel(os.path.join(path,(
                mco + ' Cap Leak ' + month + ' - ' + year + '.xlsx')), sheet_name=checklist.pop(-1))                #try to find sheet with name specified
            print('Beginning Refactor on: ' + mco + '\n')
        except FileNotFoundError:                                                                                   #catch error if not found and notify user on screen
            print('Failed to find raw file for: ' + mco + '\n\n\n')
            return
        for i in checklist:                                                                                         #find columns that are designated to be concatinated
            if len(i.split(',')) > 1:
                q = i.split(',')                                                                                    #split on diliminator ',' and check length
                df[str(q[0])] = df[str(q[0])] + ', ' + df[str(q[1])]                                                #concat two columns
                item = str(q.pop(0))                                                                                #update lists and mapping with combined column's new name
                columns_remove[columns_remove.index(i)] = item
                checklist[checklist.index(i)] = item
                self.mapping[mco][item] = self.mapping[mco].pop(i)
        df = df[df.columns.intersection(columns_remove)]                                                            #intersect the columns to remove non-used data
        df['MCOLong'] = list(columns_remove)[0]                                                                     #add MCO static columns to checklists
        df['MCO'] = list(columns_remove)[1]
        expect_len = len(columns_remove) - 1                                                                        #find the expected number of columns (minus the column regarding sheet number)
        print('\n' + str(len(df.columns)) + ' Columns in dataframe out of ' + str(expect_len) + ' Expected columns')#Checking for missing columns by comparing columns to expected columns
        if len(df.columns) != expect_len:                                                                           #find missing columns if applicable and print
            print('Missing Columns: ' + str([c for c in checklist[2:] if
                                             c not in df.columns or not 'MCO' or not 'MCOLong']))
        return df


main = CapLeak()
