import pandas as pd
file_path='物語目録.xlsx'
df=pd.read_excel(file_path)

df.sort_index(axis=1, inplace=True)
with pd.ExcelWriter('物語目録.xlsx',mode='a') as writer:
    df.to_excel(writer,sheet_name='並べ替え') 
    