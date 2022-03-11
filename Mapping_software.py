import tkinter as tk
import pyodbc as db


class Gui:

    def __init__(self):
        self.window = tk.Tk()
        self.window.geometry('400x600')
        self.window.title('Excel Column Mapper Version | 0.1')
        self.base()
        conn_string = 'Driver=SQL Server;' \
                      'Server=tp-bisql-02;' \
                      'Database=Finance;' \
                      'Trusted_Connection=yes;'  # DB connection string
        self.cnxn = db.connect(conn_string)
        self.cursor = self.cnxn.cursor()

    def base(self):
        self.baseframe = tk.Frame(master=self.window)
        self.baseframe.place(relwidth=1, relheight=1)
        self.newmapbutton = tk.Button(master=self.baseframe, text='New Map', command=self.get_table_columns)
        self.newmapbutton.place(width=100, height=30, relx=.1, rely=.1)
        self.existingmapbutton = tk.Button(self.baseframe, text='Existing Map')
        self.existingmapbutton.place(width=100, height=30, relx=.1, rely=.2)
        # self.concatwbbutton = tk.Button(self.baseframe, text='Concatenate Spreadsheets')
        # self.concatwbbutton.place(width=150, height=30, relx=.1, rely=.85)

    def get_dbtables(self):
        self.cursor.execute('SELECT TABLE_NAME from [Finance].INFORMATION_SCHEMA.TABLES;')
        dblist = list(self.cursor.fetchall())
        dblist = [i[0] for i in dblist]
        return dblist

    def get_table_columns(self):
        table_name = 'Fact_MCOCumulativeRosterCHPW'
        self.cursor.execute('SELECT COLUMN_NAME, ORDINAL_POSITION, IS_NULLABLE, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = ?', table_name)
        field_list = self.cursor.fetchall()
        y = .25
        for i in field_list:
            x = .1
            index = 0
            for q in i:
                if index == 1:
                    x += .1
                    index += 1
                    continue
                else:
                    label = tk.Label(self.baseframe, text=q)
                    label.place(relx=x, rely=y)
                    x += .15
                    index += 1
            y += .025










if __name__ == '__main__':
    gui = Gui()
    gui.window.mainloop()