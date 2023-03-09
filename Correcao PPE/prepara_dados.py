import pandas as pd
notas = pd.read_excel('notas.xlsx')

notas.columns = ['notas']

notas.dropna(inplace=True)
notas.reset_index(drop=True, inplace=True)

notas.head()

notas.shape

selecao = []
for linha in notas['notas']:
  if len(linha) <= 60:
    selecao.append(linha)

cont = 0
selecao_final = []
for item in selecao:
  if 'CNJ:' in item:
    selecao[cont] = item.split(': ')[1]
    selecao[cont] = selecao[cont].split(')')[0]
    selecao_final.append(selecao[cont])
    print(f'Primeiro: {item}')
  elif 'CNJ' in item and len(item) <=30:
    selecao[cont] = item.split('(')[0]
    selecao_final.append(selecao[cont])
    print(f'Segundo: {item}')
  elif 'CNJ' in item and len(item) > 30:
    selecao[cont] = item.split(' ')[2]
    selecao[cont] = selecao[cont].split(')')[0]
    selecao_final.append(selecao[cont])
    print(f'Terceiro: {item}')
  else:
    print('Caso fora do PPE')
  cont +=1


cont = 0
processo_sem_format = ''
for processo in selecao_final:
  while len(selecao_final[cont]) < 25:
    selecao_final[cont] = '0' + selecao_final[cont]
  processo_formatado = selecao_final[cont]
  selecao_final[cont] = ''.join(selecao_final[cont].split('.'))
  selecao_final[cont] = ''.join(selecao_final[cont].split('-'))
  selecao_final[cont] = (selecao_final[cont], processo_formatado)
  cont +=1
selecao_final

len(selecao_final)

dados_preparados = pd.DataFrame(selecao_final, index=None, columns=['sem_formatacao', 'formatado'])
dados_preparados.head(100)

dados_preparados.to_excel('Resultado.xlsx', engine='openpyxl')