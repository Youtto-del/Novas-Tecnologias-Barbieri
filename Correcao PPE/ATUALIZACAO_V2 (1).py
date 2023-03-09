from time import sleep
from selenium import webdriver
import pandas as pd
import ctypes
from selenium.common.exceptions import NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By



#DADOS
login = '01401116035'
senha = 'sumerios0394'

### INPUT DE DADOS      ###
df = pd.read_excel('Resultado.xlsx')
total_linhas=len(df.index)
print(df)
print(f'linhas: {total_linhas}')

###   NAVEGADOR   ###
navegador=webdriver.Chrome(ChromeDriverManager().install())
navegador.implicitly_wait(15)

url = 'https://ppe.tjrs.jus.br/ppe/signin'
navegador.implicitly_wait(10)
###   LOGIN   ###
navegador.get(url)      #ABRE NAVEGADOR
sleep(2)
navegador.find_element(by=By.XPATH, value='//*[@id="avisos-signin"]/div/div[3]/p-footer/button/span[2]').click()
sleep(1)
navegador.find_element(by=By.XPATH, value='//*[@id="usuario"]').send_keys(login)       #CAMPO LOGIN
sleep(0.5)
navegador.find_element(by=By.XPATH, value='//*[@id="password"]').send_keys(senha)     #CAMPO SENHA
sleep(0.5)
navegador.find_element(by=By.XPATH, value='//*[@id="container-principal"]/app-signin/div[3]/div[2]/div/form/button').click()     #BOTAO ENTRAR
sleep(1.5)
navegador.find_element(by=By.XPATH, value='//*[@id="navBar"]/button/span').click()
sleep(0.5)
navegador.find_element(by=By.XPATH, value='//*[@id="supportedContentDropdownProcesso"]/span').click()
sleep(0.5)
navegador.find_element(by=By.XPATH, value='//*[@id="processo"]/div/a/span').click()
sleep(2)

###   LISTAS    ###
cj_num_antigo = []
cj_num_originario =[]
cj_processo = []
cj_processoFormatado = []

###   PESQUISA PROCESSOS PARA CADA LINHA DO EXCEL   ###
for i in range(0, total_linhas):
  processo = df['sem_formatacao'].iloc[i]
  processoFormatado = df['formatado'].iloc[i]
  processo = str(processo)
  print(processo)
  print(f'etapa {i+1}/{total_linhas}')
  sleep(1)

  navegador.find_element(by=By.XPATH, value='//*[@id="numeroProcesso"]').send_keys(processo)      #ENVIA NUMERO PROCESSO PARA CAMPO PESQUISA
  sleep(1.5)
  navegador.find_element(by=By.XPATH, value='//*[@id="bt_pesquisar"]').click()    #BOTÃO PESQUISAR
  sleep(2)
  tentativas = 1
  while 1:
    try:
      navegador.find_element(by=By.XPATH, value='//*[@id="processos-grid"]/div/div/table/tbody/tr').click()   #ACESSA RESULTADO
      break
    except:
      tentativas += 1
      navegador.find_element(by=By.XPATH, value='//*[@id="numeroProcesso"]').send_keys(processo)      #ENVIA NUMERO PROCESSO PARA CAMPO PESQUISA
      navegador.find_element(by=By.XPATH, value='//*[@id="bt_pesquisar"]').click()    #BOTÃO PESQUISAR
      sleep(2)
  sleep(10)
  print(f"o robô procurou {tentativas} vezes.")
  #COLETA NUMERO ANTIGO E NUMERO ORIGINÁRIO
  try:
    num_antigo=navegador.find_element(by=By.XPATH, value='//*[@id="campoDetalhesProcesso"]/div[1]/div[2]/span').text
  except NoSuchElementException:
    num_antigo='Sem numero antigo'

  try:
    num_originario=navegador.find_element(by=By.XPATH, value='//*[@id="campoDetalhesProcesso"]/div[6]/div[2]/ppe-detalhes-processo-processos-vinculados/div[2]/p-datatable/div/div[1]/div/div[2]/div/table/tbody/tr/td[3]/span/td').text
  except NoSuchElementException:
    num_originario='Sem dados do processo'

  #APPENDS
  cj_processo.append(processo)    #ARMAZENA NA LISTA O NUMERO DO PROCESSO
  cj_num_antigo.append(num_antigo)     #ARMAZENA NA LISTA O NUMERO ANTIGO
  cj_num_originario.append(num_originario)     #ARMAZENA NA LISTA O NUMERO DO PROCESSO ORIGINÁRIO
  cj_processoFormatado.append(processoFormatado)  #ARMAZENA NA LISTA O NUMERO DO PROCESSO FORMATADO

  
  sleep(4)
  navegador.find_element(by=By.XPATH, value='//*[@id="navBar"]/button').click()
  sleep(0.5)
  navegador.find_element(by=By.XPATH, value='//*[@id="supportedContentDropdownProcesso"]/span').click()
  sleep(0.5)
  navegador.find_element(by=By.XPATH, value='//*[@id="processo"]/div/a/span').click()
  sleep(1)


#navegador.quit()

df={'Principal':cj_processo, 'Número antigo':cj_num_antigo, 'Originário':cj_num_originario, 'Processo Formatado': cj_processoFormatado}      #ARMAZENA OS VALORES EM UMA MATRIZ
df1=pd.DataFrame(df)        #CRIA O DATAFRAME
df1.to_excel('./PROCESSOS ATUALIZADOS.xlsx')

navegador.quit()
ctypes.windll.user32.MessageBoxW(0, "Robo notas!", "Concluído!", 1)
#CONTROLE
print(df1)