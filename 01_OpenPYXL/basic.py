import pandas
import pandas as pd

df = pd.read_csv('data.xlsx')
# df = pd.read_csv('data.xlsx', index_col='id')
print(df.head())