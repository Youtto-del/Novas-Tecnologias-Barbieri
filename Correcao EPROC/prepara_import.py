import pandas as pd
from openpyxl import load_workbook

# Carrega o modelo
modelo = 'Lista de notas.xlsx'
planilha = load_workbook(modelo)

# Leitura da aba
aba = 'importacao'
writer = pd.ExcelWriter(modelo, engine='openpyxl', mode='a')

writer.book = planilha
df = pd.read_excel(planilha, sheet_name=aba)

# seleção de células a partir de critérios
filtro1 = df['Status 1'] == 'Migrado'
filtro2 = df['Status 1'].str.contains('Digitalizado')

# aplicação dos filtros
df_filtrado = df.loc[filtro1 & filtro2]

# escrita dos dados filtrados em uma nova aba
nome_nova_aba = 'resultados'
df_filtrado.to_excel(writer, sheet_name=nome_nova_aba, index=False)

# salva as alterações na planilha
writer.save()
