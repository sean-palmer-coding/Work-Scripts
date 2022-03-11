import pandas as pd
import os
import xlrd.biffh
from sqlalchemy import create_engine
import colorama
import urllib
import datetime
import xlwings as xw

colorama.init(autoreset=True)


class UniversalMapping:

    def __init__(self):
        self.path = "H:\\Finance\\MANAGED CARE\\Shared\\BI Upload\\Universal Mapping\\"
        self.archive_base_path = "H:\\Finance\\MANAGED CARE\\Shared\\BI Upload\\MCO Data Archives\\"
        for f in os.listdir(self.path):
            print(str(os.listdir(self.path).index(f)) + ") " + f)
        while True:
            action = input("Which action would you like to perform?")
            try:
                if int(action) in [os.listdir(self.path).index(x) for x in os.listdir(self.path)]:
                    self.path = os.path.join(self.path, os.listdir(self.path)[int(action)])
                    print("\nAction accepted, mapping...\n")
                    break
                else:
                    print("Your input was not understood")
            except ValueError:
                print("Your input was not understood")
                continue
        self.mapping = pd.read_excel(os.path.join(self.path, "Mapping File.xlsx"), sheet_name="Mapping").set_index(
            'Upload_Field')
        self.meta = pd.read_excel(os.path.join(self.path, "Mapping File.xlsx"), sheet_name="Meta").set_index('Meta')
        self.date_control = pd.read_excel(os.path.join(self.path, "Mapping File.xlsx"),
                                          sheet_name="Date Input").set_index("Date Control")
        if len(str(self.date_control.loc["Month", "Value"])) < 2:
            self.date_control.loc["Month", "Value"] = "0" + str(self.date_control.loc["Month", "Value"])
        self.df_dict = {}
        self.completed_df = []
        self.connection_str = urllib.parse.quote_plus('Driver={ODBC Driver 17 for SQL Server};'
                                                      'Server=tp-bisql-02;'
                                                      'Database=Finance;'
                                                      'Trusted_Connection=yes;')
        self.engine = create_engine('mssql+pyodbc:///?odbc_connect={}'.format(self.connection_str),
                                    fast_executemany=True)
        self.main()

    def main(self):
        self.missing_file = []
        for m in self.meta.columns:
            missing_flag = 1
            if not self.meta.loc["Import?", m]:
                print("Not imported: " + m)
                continue
            p2 = self.meta.loc["File Path", m]
            for i in os.listdir(os.path.join(self.path, p2)):
                if str(self.date_control.loc["Month", "Value"]) in i.split("-")[-2] and str(self.date_control.loc["Year", "Value"]) in i.split("-")[-1]:
                    print("Reading file into memory: " + m + "\nFile Name: " + i + "\n")
                    missing_flag = 0
                    try:
                        self.df_dict[m] = pd.read_excel(
                            os.path.join(self.path, p2, i),
                            sheet_name=self.meta.loc["sheet_number", m],
                            header=self.meta.loc["header", m],
                            na_values=["#N/A", "N/A N/A", "#NA", "-1.#IND", "-1.#QNAN", "-NaN", "-nan", "1.#IND", "1.#QNAN", "<NA>", "N/A", "NULL", "NaN", "n/a", "nan", "null", ""],
                            keep_default_na=False,
                            na_filter=True,
                            dtype=object
                        )
                    except OverflowError:
                        print(colorama.Fore.LIGHTRED_EX + "An error occurred when parsing " + m + " data. This is typically caused by column type.\n"
                                "look for columns that show ###### no matter how large they are stretched.")
                        self.__init__()
                    except xlrd.biffh.XLRDError:
                        print(m + " is a protected workbook, continuing loading into memory\n")
                        wb = xw.Book(os.path.join(self.path, p2, i))
                        sheet = wb.sheets[self.meta.loc["sheet_number", m]]
                        self.df_dict[m] = sheet.range('A1').options(pd.DataFrame,
                             header=1,
                             index=False,
                             expand='table').value
                        xl = xw.apps.active.api
                        xl.Quit()
            if missing_flag:
                print("File not found: " + m)
                self.missing_file.append(m)
        if len(self.missing_file) > 0:
            print(colorama.Fore.RED + "The following files were not found during import. "
                                      "Please place files in the correct folder and hit enter to re-run")
            for x in self.missing_file:
                print(x)
            print("\nRemember, files must contain the month and year at the end of the file path, seperated by a hyphen\n"
                  " and in the designated folder. Example: This_is_a_MCO_Roster_MM-YYYY")
            input()
            self.main()
        for mco, data in self.df_dict.items():
            self.completed_df.append(self.map(mco, data))
        flag = 0
        for mco, data in self.df_dict.items():
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
            print("Here's a summary of the data: \n")
            if len(self.df_dict.keys()) == 0:
                print("No dataframes in memory. Check date on mapping")
            for mco, df in self.df_dict.items():
                print("Summary for " + mco + " Dataframe: \n\n<<<<<<<<<< " + mco + " >>>>>>>>>>\n")
                df.info(show_counts=True)
                print()
            print(colorama.Fore.GREEN + "The mapping was completed without error\n")
            print(colorama.Fore.RED + "Please read summary stats above before continuing\n")
            print(colorama.Fore.BLUE + "Your date from mapping is: " + str(self.date_control.loc["Month", "Value"]) + "/1/" + str(self.date_control.loc["Year", "Value"]))
            mapping_date = datetime.datetime.strptime(str(self.date_control.loc["Month", "Value"]) + "/1/" + str(self.date_control.loc["Year", "Value"]), "%d/%m/%Y")
            print(mapping_date.month != datetime.datetime.now().month)
            print(mapping_date.year != datetime.datetime.now().year)
            if mapping_date.month != datetime.datetime.now().month or mapping_date.year != datetime.datetime.now().year:
                proceed = input(colorama.Fore.RED + "Your file date is not in this month. Do you wish to proceed? (y/n)")
                if proceed.lower() == 'y':
                    pass
                else:
                    print(colorama.Fore.RED + "\nAction Aborted")
                    input("<<<< Press enter to restart >>>>")
                    self.main()
            print("Select from the following options:")
            print("1) Export to Excel Workbook (for debugging purposes)")
            print("2) Insert into SQL Table")
            print("3) Export Excel Workbook to designated archive folder and Insert into SQL Table")
            choice = input("> ")
            if choice == "1":
                print("Exporting Debug Excel File to: " + os.path.join(self.path, "Output/excel_output.xlsx"))
                pd.concat([x for x in self.df_dict.values()]).to_excel(
                    os.path.join(self.path, "Output/excel_output.xlsx"), index=False)
                print("\ncomplete")
            if choice == "3":
                for mco, df in self.df_dict.items():
                    if not os.path.exists(os.path.join(self.archive_base_path, self.meta.loc["Archive Path", mco])):
                        print("Archival Directory not found, creating directory..")
                        os.makedirs(os.path.join(self.archive_base_path, self.meta.loc["Archive Path", mco]))
                        print("Directory created at: " + os.path.join(self.archive_base_path, self.meta.loc["Archive Path", mco]))
                    print("\nWriting file to archive directory..")
                    print("Archive directory: " + colorama.Fore.YELLOW + os.path.join(self.archive_base_path, self.meta.loc["Archive Path", mco]))
                    print("File Name: " + colorama.Fore.BLUE + mco + " " + str(self.date_control.loc["Month", "Value"]) + "-" + str(self.date_control.loc["Year", "Value"]) + ".xlsx")
                    df.to_excel(
                        os.path.join(
                            self.archive_base_path, self.meta.loc["Archive Path", mco],
                            mco + " " + str(self.date_control.loc["Month", "Value"]) + "-" + str(self.date_control.loc["Year", "Value"]) + ".xlsx"),
                        index=False
                    )
                    print("File writing complete")
                print("\nAll files were archived")
            if choice == "2" or choice == "3":
                for mco in self.df_dict.keys():
                    print(mco + ": " + self.meta.loc["Database Table", mco])
                    print("\n     Method to be used: " + self.meta.loc["Method", mco] + "\n")
                input("<<<< Press Enter to Proceed with Insert >>>>")
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
                for x in [x for x in self.meta.loc["Database Table"].unique()]:
                    df_list = {i: self.df_dict[i] for i in self.meta.T[(self.meta.T["Database Table"] == x) & (self.meta.T["Import?"] == 1)].index}
                    if "replace" in [self.meta.loc["Method", i] for i in df_list.keys()]:
                        method = "replace"
                    elif "append" in [self.meta.loc["Method", i] for i in df_list.keys()]:
                        method = "append"
                    else:
                        print("Check database insert method in mapping (can be either 'append' or 'replace'")
                        self.main()
                    print("\nTable {}, Executing Insert Statement...".format(x))
                    df_tosql = pd.concat(df_list.values())
                    force_dt = self.mapping[~self.mapping["DataType"].isnull()]["DataType"].to_dict()
                    for col, dt in force_dt.items():
                        if dt == "datetime":
                            print(colorama.Fore.YELLOW + "Converted column " + col + " to " + dt)
                            df_tosql[col] = pd.to_datetime(df_tosql[col])
                    chunksize = 2097 // len(df_tosql.columns)
                    self.insert_sql(x, df_tosql, method, chunksize)
                    print("\nComplete")

    def insert_sql(self, table, df, method, cs):
        df.to_sql(table, self.engine, if_exists=method, index=False, schema="dbo")
        print(colorama.Fore.GREEN + "\n" + str(len(df)) + " Rows inserted into " + table + " successfully")

    def map(self, mco, df):
        """
        This is the main mapping method that takes the raw dataframe and maps it to the mapping provided.

        :param mco: Name of the MCO
        :param df: MCO's dataframe from their Workbook sent via sftp or downloaded through web portal
        :return: returns a mapped dataframe
        """
        print(colorama.Fore.BLUE + "\nRemapping Columns: " + mco + "\n")
        mapping = self.mapping[mco]
        mapping = mapping.dropna()
        for i in mapping.index:
            while True:
                try:
                    print(colorama.Fore.YELLOW + "Mapping column: " + i)
                    if str(mapping.loc[i])[-1] == "*":
                        df[i] = str(mapping.loc[i]).split("*")[0]
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
                        df[i] = df[instructions[0]].apply(lambda x: x.split(instructions[1])[int(instructions[2])])
                    else:
                        df[i] = df[mapping.loc[i]]
                except Exception:
                    print(colorama.Fore.RED + "\nAn error occured: \n" + colorama.Fore.BLUE + "MCO: " + mco + "\nColumn: "
                          + i + "\nMapped Column: " + mapping.loc[i])
                    input("Hit enter to continue after fixing mapping (Data will be reloaded into memory)")
                    self.mapping = pd.read_excel(os.path.join(self.path, "Mapping File.xlsx"),
                                                 sheet_name="Mapping").set_index(
                        'Upload_Field')
                    self.main()
                finally:
                    break
        df = df.loc[:, ~df.columns.duplicated()]
        dropped_footer = 0
        try:
            df = df[mapping.index]
        except KeyError as e:
            print("An exception has occured, check your mapping for errors\n\n" + str(e))
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
                    print(colorama.Fore.RED + "ALERT!! SEE DATA INTEGRITY MESSAGE BELOW")
                    print("Auto dropping potential total row from bottom of dataframe; MCO: " + mco + "\n")
                break
        self.df_dict[mco] = df


if __name__ == "__main__":
    main = UniversalMapping()
