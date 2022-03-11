import os
import pandas as pd

basepath = "2021 Premera Enrollment"
mlist = []


def main():
    first_flag = None
    with os.scandir(basepath) as entries:
        for entry in entries:
            if entry.is_dir():
                with os.scandir(os.path.join(basepath, entry.name)) as innerentries:
                    for i in innerentries:
                        if entry.is_file:
                            mlist.append(builddf(i.path))
    final = pd.concat(mlist, axis=0)
    writer = pd.ExcelWriter(os.path.join(basepath, 'Premera Enrollment 07-20_07-21.xlsx'), engine='xlsxwriter')
    final.to_excel(writer, index=False)
    writer.save()

def builddf(path):
    if path.split('.')[1] == 'csv':
        return pd.read_csv(path)
    else:
        return pd.read_excel(path)


main()