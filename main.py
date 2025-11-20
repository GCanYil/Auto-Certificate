import pandas as pd

df = pd.read_excel("data/veri.xlsx")
df["Tarih"] = df["Tarih"].dt.strftime("%d.%m.%y")
rows_to_list = df.values.tolist()

