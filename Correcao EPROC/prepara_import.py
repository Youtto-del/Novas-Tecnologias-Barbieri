import pandas as pd
from openpyxl import Workbook

# Carrega o modelo
modelo = 'Lista de notas.xlsx'
planilha = Workbook()
sheet = planilha.active


df = pd.read_excel(modelo)

# seleção de células a partir de critérios
filtro_status1_1 = df[df['Status 1'] == 'Migrado']
filtro_status1_2 = df[df['Status 1'].str.contains('Digitalizado')]

filtro_status2_1 = df[df['Status 2'] == 'Migrado']
filtro_status2_2 = df[df['Status 2'].str.contains('Digitalizado')]

filtro_status3_1 = df[df['Status 3'] == 'Migrado']
filtro_status3_2 = df[df['Status 3'].str.contains('Digitalizado')]

frames = [filtro_status1_1, filtro_status1_2, filtro_status2_1, filtro_status2_2, filtro_status3_1, filtro_status3_2]

# aplicação dos filtros
df_filtrado = pd.concat(frames)


# escrita dos dados filtrados em uma nova aba
df_filtrado.to_excel('Filtrado.xlsx', index=False)
