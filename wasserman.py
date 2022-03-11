import pandas as pd

path = "C:\\Users\\SPalmer\\Documents\\wasserman.xlsx"

wasserman = pd.read_excel(path)

print(wasserman[wasserman["Code\n-----"] != "Code\n-----"].rename(
    columns={
        k: v for (k, v) in zip(
            wasserman.columns, [i.split("\n")[0] for i in wasserman.columns]
        )
    }
))

