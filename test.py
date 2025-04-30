import pandas as pd

df = pd.read_excel("Dillard Street Durham NC/DURHAM NC - Mid June.xlsx", skiprows=8, header=0)
df = df[["CUSTOMER PO#", "OI#", "SKU", "PRICE", "TALLY", "CHECK POD"]]
df = df[df["OI#"].notna()]
print(df)