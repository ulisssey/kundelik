import pandas as pd


pd.set_option('display.max_rows', None)
pd.set_option('display.max_columns', None)
df = pd.read_excel(r"C:\Users\erasy\Downloads\original.xls")

col1 = df.columns[0]
class_num = col1.split('Класс: ')[1]

new_df = df[df['Unnamed: 1'] == 1].loc[:, 'Unnamed: 1']
rows = list(new_df.index)
rows.append(df.shape[0])

i = 0
while i < len(rows):
    start = rows[i]
    if start == rows[-1]:
        break
    end = rows[i+1]
    print(df.iloc[start:end-2])
    i += 1