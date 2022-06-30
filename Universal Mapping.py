import time
import xlrd
import pandas as pd
import os
from shutil import copy2
import sqlalchemy.exc
import xlrd.biffh
from sqlalchemy import create_engine
import colorama
import urllib
import datetime
import numpy as np
import xlwings as xw

colorama.init(autoreset=True)


class UniversalMapping:
    """
    This universal mapping program was created in order to map any excel workbook to any database table.
    The program will walk you through the process, and provides robust error catching for known errors in current processes
    The primary working directory that this program is designed to work in can be seen or set in the self.path variable in the class initializer.
    Variable definitions will be provided via inline comments.

    Each process is contained in a folder that is in the base directory (self.path), and you can add a new process by adding a folder that follows the same folder structure
    example:
    Universal Mapping/
    ├──Cap Leak/
    ├──Cap Payments/
    ├──Cumulative Roster/
    ├──Roster/
    ├──State Payments/
    ├──<Example Process>/
        ├──<Year**>/
            ├──<Workbook Type/ MCO**>/
        ├──Output/
        ├──Mapping File.xlsx
    **This structure can be adjusted on the [Meta] tab of the mapping file, and is just an example. Files without '**' need to be in the relative position shown and exist.

    All the current mappings follow this format.
    There are three tabs on each mapping. The first is [Mapping] which contains the mapping for the workbook.
    Second is the [Meta] tab, which contains metadata regarding the mapping, including things like sheet number (to collect data from)
    File path where the data can be located, database table to be inserted into, the method used to insert, header row, and archive paths.
    Finally, there is the [Date Input] tab, which contains the month and year that you would like to run the process for. This mapping assumes that
    the mapping that you would like to run follows these conventions, format, and dimensions.

    The mapping program also allows you to concatenate, split, hardcode values, and force datatype conversions through a simple syntax in the mapping file
    This makes it so you do not have to hardcode processes in python.

    Concatenation (using the '+' operator):
        To concatenate two columns, follow the format of <column name>+<column name>. Other examples include:
            <column name>+<column name>+<column name>
            <column name>+,+<column name>
        You can choose to concatenate a string in between the columns if you would like, for example:
            Member_Last_Name+,+Member_First_Name
        Would take the columns Member_Last_Name and Member_First_Name and concatenate them together with a comma in the middle like Smith,John

    Splits (using the '|' operator):
        To split a column, follow the format of <column name>|<character to split on>|<index>. Example:
            Member Name|,|0
        Which would yield from the value 'Smith, John' --> 'Smith'
            Member Name|,|1
        Which would yield from the value 'Doe, Jane' --> 'Jane'

    Fixed Values (using the '*' operator):
        To create a fixed value in a column, all that is needed is to add the '*' character to the end of the fixed value you would like to add
        on the end of the string you'd like to use as a fixed value. for example:
            Molina*
        Which would yield a column that every value for every row contained the string 'Molina'. You can even use a calculated column in excel
        and add this to the end of the calculated\concatenated column to create a value that is fixed in the final table. example:
            ='Date Input'!$B$2&"/1/"&'Date Input'!$B$3&"*"
        Which would yield a fixed date based on the input on the [Date Input] tab of the workbook.

    Forced Datatyping (For now, just dates, datetimes, and strings):
        On every mapping, there is a column that is called DataType which is the first column on every mapping. This column can be used to force the datatype of the
        finished column to a datetime (with time stamp) or just a date (without time stamp)
        For timestamped datetimes, type 'datetime' in that column. For dates, type 'date' into that column. For strings,
        type 'string'.

    Error Checking:
        There are multiple areas for error checking, although not all errors will be caught. Be sure to read the error carefully and try to use the solutions
        provided on errors where a suggested solution is provided. If there is an uncaught error, save the error for debugging later.

    Mapping:
        Each folder has it's own mapping which you will need to visit to update the year. As long as this file is not auto saved enabled, you may leave the file
        open when you run the program, and make changes to the file and save them as long as the program is not actively trying to access the map while you're saving.
        Inside the universal mapping folder there is a word document that further explains the process of mapping a column and using the mapping workbook.

    For more detailed documentation, see WordDoc in BIUpload folder Universal Mapping Documentation.docx
    """

    def __init__(self):
        self.path = "H:\\Finance\\MANAGED CARE\\Shared\\BI Upload\\Universal Mapping\\"                                 #Base path for all processes
        self.archive_base_path = "H:\\Finance\\MANAGED CARE\\Shared\\BI Upload\\MCO Data Archives\\"                    #Base path that the archive directory join to based off of mapping
        for f in os.listdir(self.path):                                                                                 #prints options based on folders in the first level of the base directory
            print(str(os.listdir(self.path).index(f)) + ") " + f)
        while True:                                                                                                     #Menu for selecting the process you want to perform
            action = input("Which action would you like to perform?")
            try:
                if int(action) in [os.listdir(self.path).index(x) for x in os.listdir(self.path)]:
                    self.path = os.path.join(self.path, os.listdir(self.path)[int(action)])
                    print("\nAction accepted, mapping...\n")
                    break
                else:
                    print("Your input was not understood")                                                              #Error checking on user input
            except ValueError:
                print("Your input was not understood")
                continue
        self.mapping = pd.read_excel(os.path.join(self.path, "Mapping File.xlsx"), sheet_name="Mapping").set_index(     #reading [Mapping] tab into memory
            'Upload_Field')
        self.meta = pd.read_excel(os.path.join(self.path, "Mapping File.xlsx"), sheet_name="Meta").set_index('Meta')    #reading [Meta} tab into memory
        self.date_control = pd.read_excel(os.path.join(self.path, "Mapping File.xlsx"),                                 #reading [Date Input] tab into memory
                                          sheet_name="Date Input").set_index("Date Control")
        if len(str(self.date_control.loc["Month", "Value"])) < 2:                                                       #correctly formating date from date control tab for error checking later
            self.date_control.loc["Month", "Value"] = "0" + str(self.date_control.loc["Month", "Value"])
        self.df_dict = {}                                                                                               #instantiating variables to be used later
        self.completed_df = []
        self.connection_str = urllib.parse.quote_plus('Driver={ODBC Driver 17 for SQL Server};'                         #connection string, read more about this here: https://docs.microsoft.com/en-us/sql/connect/python/pyodbc/step-3-proof-of-concept-connecting-to-sql-using-pyodbc?view=sql-server-ver15
                                                      'Server=tp-bisql-02;'
                                                      'Database=Finance;'
                                                      'Trusted_Connection=yes;')
        self.engine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(self.connection_str),                     #creates the SQL server engine, more info here: https://docs.sqlalchemy.org/en/14/core/connections.html
                                    fast_executemany=True)
        self.missing_file = []                                                                                          #instantiate more variables for use later
        self.file_list = {}
        self.df_tosql = None
        self.main()                                                                                                     #call the main function

    def main(self):
        """
        Bulk of program that contains menus, error catching, reading and writing data into and out of memory.

        :return: None;
        """

        for m in self.meta.columns:                                                                                     #begin iteration through columns on [Meta] tab
            missing_flag = 1
            if not self.meta.loc["Import?", m]:                                                                         #check to see if file should be imported, print if not imported
                print("Not imported: " + m)
                continue
            p2 = self.meta.loc["File Path", m]                                                                          #if file needs to be imported, look at the filepath cell related and read file into memory
            for i in os.listdir(os.path.join(self.path, p2)):
                try:
                    if str(self.date_control.loc["Month", "Value"]) in i.split("-")[-2] and str(                        #look for the month and year seperated by a '-' in the file name
                            self.date_control.loc["Year", "Value"]) in i.split("-")[-1]:
                        print("Reading file into memory: " + m + "\nFile Name: " + i + "\n")
                        self.file_list[m] = (os.path.join(self.path, p2, i))                                            #add file to a file list used to copy to archives later
                        missing_flag = 0
                        try:
                            self.df_dict[m] = pd.read_excel(                                                            #read the excel data from the file specified into memory and add the a dictionary for easy access later
                                os.path.join(self.path, p2, i),
                                sheet_name=self.meta.loc["sheet_number", m],
                                header=self.meta.loc["header", m],
                                na_values=[
                                    "#N/A", "N/A N/A", "#NA", "-1.#IND", "-1.#QNAN", "-NaN", "-nan", "1.#IND",
                                    "1.#QNAN", "<NA>", "N/A", "NULL", "NaN", "n/a", "nan", "null", ""
                                ],
                                keep_default_na=False,
                                na_filter=True,
                                dtype=object
                            )
                        except OverflowError:                                                                           #error checking for a specific error that is caused by incorrect style/datatype applied in excel
                            print(colorama.Fore.LIGHTRED_EX + "An error occurred when parsing " + m + " data. This is"
                                    " typically caused by column type.\n"
                                    "look for columns that show ###### no matter how large they are stretched.")
                            self.__init__()
                        except xlrd.biffh.XLRDError:
                            print(m + " is a protected workbook, continuing loading into memory\n")                     #error catch for protected workbooks, will attempt to open file a different way and load into memory
                            wb = xw.Book(os.path.join(self.path, p2, i))
                            sheet = wb.sheets[self.meta.loc["sheet_number", m]]
                            self.df_dict[m] = sheet.range('A1').options(pd.DataFrame,
                                 header=1,
                                 index=False,
                                 expand='table').value
                            xl = xw.apps.active.api
                            xl.Quit()
                        except ValueError:
                            self.df_dict[m] = pd.read_csv(
                                os.path.join(self.path, p2, i),
                                header=self.meta.loc["header", m],
                                na_values=["#N/A", "N/A N/A", "#NA", "-1.#IND", "-1.#QNAN", "-NaN", "-nan", "1.#IND",
                                           "1.#QNAN", "<NA>", "N/A", "NULL", "NaN", "n/a", "nan", "null", ""],
                                keep_default_na=False,
                                na_filter=True,
                                dtype=object
                            )
                except IndexError:
                    print("Mapping File Month and Year: " + str(self.date_control.loc["Month", "Value"]) + "/" + str(
                        self.date_control.loc["Year", "Value"]))                                                        #a file was not named correctly, and needs to be fixed. hitting enter restarts process
                    print("One of the files in the folder does not contain a hyphen seperated date. "
                          "Please fix and run again")
                    input("<<< Press Enter to Re-run Script >>>")
                    self.__init__()
            if missing_flag:
                print("File not found: " + m)
                self.missing_file.append(m)                                                                             #if the file is missing, ads to a list
        if len(self.missing_file) > 0:
            print(colorama.Fore.RED + "The following files were not found during import. "
                                      "Please place files in the correct folder and hit enter to re-run")
            for x in self.missing_file:
                print(x)
            print("\nRemember, files must contain the month and year at the end of the file path, seperated by a hyphen\n"
                  " and in the designated folder. Example: This_is_a_MCO_Roster_MM-YYYY")                               #hitting enter re-runs script after a file that was supposed to be imported was not found
            input("<<< Press Enter to Re-run Script >>>")
            self.main()
        for mco, data in self.df_dict.items():                                                                          #call mapping method and remap the dataframes
            self.completed_df.append(self.map(mco, data))
        flag = 0
        for mco, data in self.df_dict.items():                                                                          #if an error is thrown, mapping will return a str, which means it failed
            if type(data) == str:
                flag = 1
        if flag:
            while True:
                print("There was an error, check mapping an rerun")
                print("Rerun? (fix mapping errors first and save workbook)")
                choice = input("(y/n)?")
                if choice.lower() == "y":
                    self.main()
                elif choice.lower() == "n":
                    return
                else:
                    print("Your Response was not understood")
        else:
            print("Here's a summary of the data: \n")                                                                   #summarizes data so you can see the row counts, columns, and datatypes
            if len(self.df_dict.keys()) == 0:
                print("No dataframes in memory. Check date on mapping")
            for mco, df in self.df_dict.items():
                print("Summary for " + mco + " Dataframe: \n\n<<<<<<<<<< " + mco + " >>>>>>>>>>\n")
                df.info(show_counts=True)
                print()
            print(colorama.Fore.GREEN + "The mapping was completed without error\n")
            print(colorama.Fore.RED + "Please read summary stats above before continuing\n")
            print(colorama.Fore.BLUE + "Your date from mapping is: " + str(
                self.date_control.loc["Month", "Value"]
            ) + "/1/" + str(self.date_control.loc["Year", "Value"]))
            mapping_date = datetime.datetime.strptime(str(self.date_control.loc["Month", "Value"]) + "/1/" + str(
                self.date_control.loc["Year", "Value"]
            ), "%m/%d/%Y")
            if mapping_date.month != datetime.datetime.now().month or mapping_date.year != datetime.datetime.now().year:
                proceed = input(
                    colorama.Fore.RED + "Your file date is not in this month. Do you wish to proceed? (y/n)"
                )
                if proceed.lower() == 'y':                                                                              #check to make sure the user is aware of the date that they entered. If it is different from current month, override is required
                    pass
                else:
                    print(colorama.Fore.RED + "\nAction Aborted")
                    input("<<<< Press enter to restart >>>>")
                    self.main()
            print("Select from the following options:")                                                                 #menu to decide what type of insert to do
            print("1) Export to Excel Workbook (for debugging purposes)")
            print("2) Insert into SQL Table")
            print("3) Export Excel Workbook to designated archive folder(s) and Insert into SQL Table")
            choice = input("> ")
            if choice == "1":
                print("Exporting Debug Excel File to: " + os.path.join(self.path, "Output/excel_output.xlsx"))
                pd.concat([x for x in self.df_dict.values()]).to_excel(
                    os.path.join(self.path, "Output/excel_output.xlsx"), index=False)                                   #exports data to the output folder for debugging
                print("\ncomplete")
            if choice == "3":
                for mco, df in self.df_dict.items():                                                                    #this creates folders if the do not exist for archiving purposes.
                    if not os.path.exists(os.path.join(self.archive_base_path, self.meta.loc["Archive Path", mco])):    #these parameters can be set in the mapping file
                        print("Archival Directory not found, creating directory..")
                        os.makedirs(os.path.join(self.archive_base_path, self.meta.loc["Archive Path", mco]))
                        print("Directory created at: " + os.path.join(
                            self.archive_base_path, self.meta.loc["Archive Path", mco])
                              )
                    if not os.path.exists(self.meta.loc["Raw Archive Path", mco]):
                        print("Raw archival directory not found, creating directory...")
                        os.makedirs(self.meta.loc["Raw Archive Path", mco])
                        print("Directory created at: " + self.meta.loc["Raw Archive Path", mco])
                    print("\nWriting file to archive directories..")
                    filename = mco + " " + str(self.date_control.loc["Month", "Value"]) + "-" + str(
                        self.date_control.loc["Year", "Value"]) + ".xlsx"
                    print("Raw archive directory: " + colorama.Fore.YELLOW + str(os.path.join(
                        self.meta.loc["Raw Archive Path", mco], filename)
                    ))
                    print("Archive directory: " + colorama.Fore.YELLOW + os.path.join(
                        self.archive_base_path, self.meta.loc["Archive Path", mco]
                    ))
                    print("File Name: " + colorama.Fore.BLUE + mco + " " +
                          str(self.date_control.loc["Month", "Value"]) + "-"
                          + str(self.date_control.loc["Year", "Value"]) + ".xlsx")
                    df.to_excel(
                        os.path.join(
                            self.archive_base_path, self.meta.loc["Archive Path", mco],
                            mco + " " + str(self.date_control.loc["Month", "Value"]) + "-" +
                            str(self.date_control.loc["Year", "Value"]) + ".xlsx"),
                        index=False
                    )
                    copy2(self.file_list[mco], self.meta.loc["Raw Archive Path", mco])
                    print("File writing/copying complete")
                print("\nAll files were archived")
            if choice == "2" or choice == "3":                                                                          #confirms user wants to truncate a table before insert
                for mco in self.df_dict.keys():
                    print(mco + ": " + self.meta.loc["Database Table", mco])
                    print("\n     Method to be used: " + self.meta.loc["Method", mco] + "\n")
                if "replace" in self.meta.loc["Method"].values:
                    print("Data in the following table will be truncated: ")
                    for t in self.meta.T[(self.meta.T["Method"] == "append")]["Database Table"].to_list():
                        print(t)
                    print("\nIf a table is seen twice, check mapping and ensure that replace is only on the FIRST"
                          " MCO to use that table.")
                    print("To confirm that you would like to truncate the table, "
                          "type TRUNCATE below, or exit the program and remap")
                    choice_trunc = input(">")
                    if choice_trunc != "TRUNCATE":
                        print("Exiting, response was not 'TRUNCATE'")
                        return
                for x in [x for x in self.meta.loc["Database Table"].unique()]:                                         #checks method for each individual table in the mapping, regardless of file and groups files by table
                    df_list = {i: self.df_dict[i] for i in self.meta.T[(self.meta.T["Database Table"] == x)
                                                                       & (self.meta.T["Import?"] == 1)].index}
                    if 1 in [self.meta.loc["Import?", i] for i in df_list.keys()]:
                        pass
                    else:
                        print("Nothing imported for table " + x)
                        print("Continuing...")
                        continue
                    if "replace" in [self.meta.loc["Method", i] for i in df_list.keys()]:                               #!!!! This method needs work!
                        method = "replace"
                    elif "append" in [self.meta.loc["Method", i] for i in df_list.keys()]:
                        method = "append"
                    else:
                        print(x)
                        print("Check database insert method in mapping (can be either 'append' or 'replace'")           #error checking for incorrect [Meta] entry
                        self.main()
                    input("<<< Insert Ready, press enter to begin SQL insert >>>")
                    print("\nTable {}, Executing Insert Statement...".format(x))
                    df_tosql = pd.concat(df_list.values())
                    df_tosql = df_tosql.applymap(lambda x: x.strip() if isinstance(x, str) else x)                      #removes all leading and trailing spaces from all str columns
                    force_dt = self.mapping[~self.mapping["DataType"].isnull()]["DataType"].to_dict()
                    for col, dt in force_dt.items():                                                                    #forces date and datetimes on columns specified
                        if dt == "datetime":
                            print(colorama.Fore.YELLOW + "Converted column " + col + " to " + dt)
                            df_tosql[col] = pd.to_datetime(df_tosql[col])
                            print(colorama.Fore.GREEN + "Success")
                        if dt == "date":
                            print(colorama.Fore.YELLOW + "Converting column " + col + " to " + dt)
                            try:
                                if str(df_tosql[col].dtypes) == "int64":
                                    df_tosql[col] = pd.to_datetime(df_tosql[col].apply(self.read_date), errors='coerce')
                                    print("Used alternate date conversion for int64 dates")
                                    df_tosql[col] = df_tosql[col].fillna(pd.to_datetime('2199-12-31'))
                                else:
                                    df_tosql[col] = pd.to_datetime(df_tosql[col], utc=False)
                                print(colorama.Fore.GREEN + "Success")
                            except ValueError as e:                                                                     #this is for dates that are larger than a 64bit number when converting
                                df_tosql[col] = df_tosql[col].apply(
                                    lambda x: x if str(x) == 'nan' else datetime.datetime.strptime(str(x), "%Y-%m-%d %H:%M:%S")
                                )
                                print("Using Alternate Date Conversion method because..")
                                print(e)
                            except KeyError:
                                print(colorama.Fore.YELLOW + "Column not found, continuing...")
                                pass

                        if dt == "string":
                            print(colorama.Fore.YELLOW + "Converting column " + col + " to " + dt)
                            df_tosql[col] = pd.Series(df_tosql[col], dtype="string")
                            print(colorama.Fore.GREEN + "Success")
                    chunksize = 2097 // len(df_tosql.columns)    #calculates chunksize
                    self.df_tosql = df_tosql
                    try:
                        self.insert_sql(x, df_tosql, method, chunksize)
                    except sqlalchemy.exc.ProgrammingError:                                                             #some inserts will through errors when using fast insert. This attempts a slow insert on error
                        print(colorama.Fore.RED + "Fast execution error. Attempting slow execution insert, "
                                                  "this can take several minutes. \n\n" 
                              "       (  )   (   )  )\n" 
                              "       ) (   )  (  ( \n" 
                              "       ( )  (    ) )\n" 
                              "       _____________\n" 
                              "      <_____________> ___\n" 
                              "      |             |/ _ \ \n" 
                              "      |               | | | \n"
                              "      |               |_| | \n"
                              "   ___|             |\___/ \n"
                              "  /    \___________/    \ \n"
                              "  \_____________________/")
                        self.engine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(self.connection_str),
                                    fast_executemany=False)
                        self.insert_sql(x, df_tosql, method, chunksize)
                    except sqlalchemy.exc.DataError:                                                                    #error catch for incorrectly formatted date times.
                        print(colorama.Fore.RED + "Invalid character value for cast specification.\n")
                        print("Check to make sure columns that have dates are forced as dates or datetimes. \n"
                              "In the past this has caused this error.")
                    print("\nComplete")
                quit()

    def read_date(self, date):
        return xlrd.xldate.xldate_as_datetime(date, 0)

    def insert_sql(self, table, df, method, cs):
        """
        This is the insert method that inserts data into the sql table
        :param table: Name of the table to be inserted into in the Finance Database
        :param df: The dataframe to be inserted into the table
        :param method: truncate or append; truncate requires override above
        :param cs: Chunksize; not used
        :return: None; prints successful insert message
        """
        start = time.time()
        print("Insertion in progress...")
        df.to_sql(table, self.engine, if_exists=method, index=False, schema="dbo")
        print(colorama.Fore.GREEN + "\n" + str(len(df)) + " Rows inserted into " + table + " successfully")
        end = time.time()
        print("Elapsed Time: " + str(int((end-start)/60)) + ":" + str(int((end-start)%60)))

    def map(self, mco, df):
        """
        This is the main mapping method that takes the raw dataframe and maps it to the mapping provided.

        :param mco: Name of the MCO
        :param df: MCO's dataframe from their Workbook sent via sftp or downloaded through web portal
        :return: returns a mapped dataframe
        """
        print(colorama.Fore.BLUE + "\nRemapping Columns: " + mco + "\n")                                                #prints column name to show which dataframe is being mapped
        mapping = self.mapping[mco]
        mapping = mapping.dropna()
        for i in mapping.index:                                                                                         #iterates through mapping
            while True:
                try:
                    print(colorama.Fore.YELLOW + "Mapping column: " + i)
                    if str(mapping.loc[i])[-1] == "*":
                        df[i] = str(mapping.loc[i]).split("*")[0]                                                       #checks for special mapping syntax characters (see Universal Mapping class docstring)
                        try:
                            df[i] = pd.to_datetime(df[i])
                        except Exception:
                            continue
                    elif "+" in str(mapping.loc[i]):
                        columns = str(mapping.loc[i]).split("+")
                        for c in columns:
                            if c not in df.columns:
                                df[c] = c
                        df[i] = df[columns].apply(lambda x: "".join(x.dropna().astype(str)), axis=1)
                    elif "|" in str(mapping.loc[i]):
                        instructions = str(mapping.loc[i]).split("|")
                        df[i] = df[instructions[0]].apply(
                            lambda x: x.split(instructions[1])[int(instructions[2])].strip() if not pd.isnull(x) else np.NaN#If cell is null than give null as return
                        )
                    else:
                        df[i] = df[mapping.loc[i]]
                except Exception as e:                                                                                  #catches exceptions on columns that did not map properly and prints which column is troublesome
                    print(colorama.Fore.RED + "\nAn error occured: \n" + colorama.Fore.BLUE + "MCO: " + mco + "\nColumn: "
                          + i + "\nMapped Column: " + mapping.loc[i])
                    print("\n" + colorama.Fore.RED + "Exception: " + str(type(e).__name__) + "\n" + str(e.args))
                    input("Hit enter to continue after fixing mapping (Data will be reloaded into memory)")
                    self.mapping = pd.read_excel(os.path.join(self.path, "Mapping File.xlsx"),
                                                 sheet_name="Mapping").set_index(                                       # re-reading [Mapping] tab into memory
                        'Upload_Field')
                    self.meta = pd.read_excel(os.path.join(self.path, "Mapping File.xlsx"),
                                              sheet_name="Meta").set_index('Meta')                                      # re-reading [Meta] tab into memory
                    self.date_control = pd.read_excel(os.path.join(self.path, "Mapping File.xlsx"),
                                                      # re-reading [Date Input] tab into memory
                                                      sheet_name="Date Input").set_index("Date Control")
                    self.main()
                finally:
                    break
        df = df.loc[:, ~df.columns.duplicated()]
        dropped_footer = 0
        try:
            df = df[mapping.index]
        except KeyError as e:
            print("An exception has occured, check your mapping for errors\n\n" + str(e))                               #missing columns from mapping or vice versa
            print(mco)
            self.df_dict[mco] = "Error Occurred, check column: " + str(e).split("'")[1]
            return
        while True:
            if df.iloc[-1, :].isnull().sum() > len(df.columns) * .6:
                dropped_footer = 1
                df.drop(index=df.index[-1], axis=0, inplace=True)
            else:
                print(len(df.columns) * .6)
                print(df.iloc[-1, :])
                if dropped_footer:
                    print(colorama.Fore.RED + "ALERT!! SEE DATA INTEGRITY MESSAGE BELOW")                               #auto drops rows that meet a certain criteria from the bottom of the dataframe that are likely total rows that should not be inserted into database
                    print("Auto dropping potential total row from bottom of dataframe; MCO: " + mco + "\n")
                break
        self.df_dict[mco] = df


if __name__ == "__main__":                                                                                              #instantiate class object
    main = UniversalMapping()
