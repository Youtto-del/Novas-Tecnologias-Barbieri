import pandas as pd
import numpy as np
from openpyxl import load_workbook

df = pd.read_excel("Lista de notas.xlsx")
df1_simplificado = df[df[['Status 1', 'Status 2', 'Status 3']].isin(['Digitalizado', 'Migrado']).any(axis=1)]

df2 = df1_simplificado[(df1_simplificado['originario_1'] != 'Não encontrado') & (df1_simplificado['Status 1'].isin(['Migrado', 'Digitalizado']))]
df2.drop(df2.iloc[:, :1], inplace=True, axis=1)
df2.drop(df2.iloc[:, 3:], inplace=True, axis=1)

df3 = df1_simplificado[(df1_simplificado['originario_2'] != 'Não encontrado') & (df1_simplificado['Status 2'].isin(['Migrado', 'Digitalizado']))]
df3.drop(df3.iloc[:, :1], inplace=True, axis=1)
df3.drop(df3.iloc[:, 1:3], inplace=True, axis=1)
df3.drop(df3.iloc[:, 3:], inplace=True, axis=1)
df4 = df1_simplificado[(df1_simplificado['originario_3'] != 'Não encontrado') & (df1_simplificado['Status 3'].isin(['Migrado', 'Digitalizado']))]
df4.to_excel('df4_teste.xlsx')
df4.drop(df4.iloc[:, :1], inplace=True, axis=1)
df4.drop(df4.iloc[:, 1:5], inplace=True, axis=1)
df4.drop(df4.iloc[:, -1:], inplace=True, axis=1)
df4.to_excel('df4_corrigido.xlsx')


df_final = pd.concat([df2, df3, df4], ignore_index=True)
print(df_final)
df_final.to_excel('df_final.xlsx')
#
# df_final = pd.DataFrame(np.concatenate((df2.values, df3.values, df4.values), axis=0))
# df_final.columns = ['Numero Principal', 'Numero Antigo', 'Pasta Desdobramento']
# df_final['Pasta'] = df_final['Pasta Desdobramento']
# df_final['Pasta'] = df_final['Pasta'].str[:-3]
# new_cols = ['Pasta', 'Pasta Desdobramento', 'Numero Antigo', 'Numero Principal']
# df_final = df_final[new_cols]
# print(df_final)

# #load workbook
# app = xw.App(visible=False)
# wb = xw.Book('Processos_correcao_EPROC_ATT_-_03102022.xlsx')
# ws = wb.sheets['Importação']
#
# #Update workbook at specified range
# ws.range('A5').options(index=False,header=False).value = df_final
#
# #Close workbook
# wb.save()
# wb.close()
# app.quit()
