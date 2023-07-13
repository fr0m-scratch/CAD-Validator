import pandas as pd

df = pd.read_excel("output.xlsx")


filtered = df[df.Text.str.match('\d+:\ 1')].index
entities = []
for row in filtered:
    if df['Graph_id'][row][9] == '1':
        diagram_number = 'FD-'+df['Text'][row].split(':')[0]
    if df['Graph_id'][row][9] == '2':
        diagram_number = 'SAMA-'+df['Text'][row].split(':')[0]
    df['diagram\nnumber'].replace(df['diagram\nnumber'][row], diagram_number, inplace=True)
print(df.head())
df.to_excel("test.xlsx")