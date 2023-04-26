from datetime import datetime
from time import sleep
import pandas as pd
from easygui import msgbox
from selenium.common import NoSuchElementException, ElementClickInterceptedException
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from pathlib import Path
# import novo_relatorio as re
from prepara_import import prepara_import


# cria o arquivo relatorios_desdobramentos
# re.relatorio_email()


def acessa_eproc():
    login = 'RS036798'
    senha = 'Barbieri5515*'
    url = 'https://eproc1g.tjrs.jus.br/eproc/'

    navegador.get(url)  # ABRE NAVEGADOR
    navegador.find_element(by=By.XPATH, value='//*[@id="txtUsuario"]').send_keys(login)  # CAMPO LOGIN
    navegador.find_element(by=By.XPATH, value='//*[@id="pwdSenha"]').send_keys(senha)  # CAMPO SENHA
    navegador.find_element(by=By.XPATH, value='//*[@id="sbmEntrar"]').click()  # BOTO ENTRAR
    msgbox("Resolva o Captcha caso apareça!")
    navegador.find_element(by=By.XPATH, value='//*[@id="tr0"]').click()  # SELECIONA PERFIL ADVOGADO


def download_planilha_comum():
    sleep(25)
    planilha = navegador.find_element(
        by=By.XPATH,
        value='//*[@id="conteudoCitacoesIntimacoesRS"]/div[2]/table/tbody/tr[7]/td[2]/a'
    )
    actions.move_to_element(planilha).click().perform()
    sleep(0.5)
    navegador.find_element(By.XPATH, value='//*[@id="conteudoCitacoesIntimacoesRS"]/div[2]/table/tbody/tr[4]/td[2]/a').click()
    sleep(2)
    janelas = navegador.window_handles
    navegador.switch_to.window(janelas[1])
    try:
        navegador.find_element(by=By.XPATH, value='//*[@id="sbmPlanilha"]').click()
    except ElementClickInterceptedException:
        print('Erro ao baixar a planilha. Usando outra alternativa...')
        button = navegador.find_element(by=By.XPATH, value='//*[@id="sbmPlanilha"]')
        navegador.execute_script("arguments[0].click();", button)

    sleep(4)
    print('Download feito')


def importa_relatorio(relatorio):
    print('Importando relatório de desdobramentos...')
    desdobramentos = pd.read_excel(relatorio)
    colunas_tabela = list(desdobramentos.columns)
    colunas_novas = list(desdobramentos.iloc[2])
    dici = {}
    for x in range(len(colunas_tabela)):
        dici[colunas_tabela[x]] = colunas_novas[x]
    desdobramentos.rename(columns=dici, inplace=True)
    desdobramentos.drop(axis=0, index=[0, 1, 2], inplace=True)
    return desdobramentos


def importa_intimacoes():
    print('Importando intimações...')
    p = Path.cwd()
    arquivo = [x for x in p.iterdir() if x.is_file() if x.stem[:16] == 'citacaoIntimacao']
    print(len(arquivo))
    print(arquivo)
    data = datetime.now().strftime("%d%m%y")
    Path.replace(arquivo[0], rf'.\log\Intimacoes_{data}.xls')
    intimacoes = pd.read_excel(rf'.\log\Intimacoes_{data}.xls')
    return intimacoes


def realiza_triagem():
    print('Realizando triagem...')
    desdobramentos = importa_relatorio('relatorio_desdobramentos.xlsx')
    intimacoes = importa_intimacoes()

    colunas_tabela = list(intimacoes.columns)
    colunas_novas = list(intimacoes.iloc[0])
    dici = {}
    for x in range(len(colunas_tabela)):
        dici[colunas_tabela[x]] = colunas_novas[x]
    intimacoes.rename(columns=dici, inplace=True)
    intimacoes.drop(axis=0, index=[0], inplace=True)
    print(intimacoes.head())

    cadastrados = []
    nao_cadastrados = []
    for intimacao in intimacoes['Processo']:
        if intimacao in set(desdobramentos['numero']):
            cadastrados.append(intimacao)
        else:
            nao_cadastrados.append(intimacao)
    return cadastrados, nao_cadastrados, desdobramentos


def coleta_originarios(cadastrados, consultas, desdobramentos):
    cj_processos = []
    cj_originarios = []
    cj_precatorios = []
    cj_originarios3 = []
    cj_datas = []
    cj_status1 = []
    cj_status2 = []
    cj_status3 = []

    item = 1
    primeiro_acesso = True
    for processo in consultas:
        print('x' * 50)
        print(processo)
        if primeiro_acesso:
            navegador.switch_to.window(navegador.window_handles[1])
            navegador.close()
            navegador.switch_to.window(navegador.window_handles[0])
            primeiro_acesso = False
        navegador.find_element(by=By.ID, value='txtNumProcessoPesquisaRapida').send_keys(processo)  # CAMPO PESQUISA
        try:
            navegador.find_element(by=By.NAME, value='btnPesquisaRapidaSubmit').click()  # BOTAO PARA PESQUISAR
        except:
            pesquisa = navegador.find_element(by=By.NAME, value='btnPesquisaRapidaSubmit')
            navegador.execute_script("arguments[0].click();", pesquisa)

        # DADOS DE CAPA
        # TENTA COLETAR PROCESSO ORIGINÁRIO
        navegador.implicitly_wait(2)
        try:
            processo_originario_1 = navegador.find_element(
                by=By.XPATH,
                value='//*[@id="tableRelacionado"]/tbody/tr/td[1]').text

            if len(processo_originario_1) >= 20:
                processo_originario_1 = processo_originario_1[0:-3]

            status1 = navegador.find_element(
                by=By.XPATH,
                value='//*[@id="tableRelacionado"]/tbody/tr/td[3]').text

        except NoSuchElementException:
            processo_originario_1 = 'Sem dados de processo originário 1'
            status1 = 'Nulo'

        # TENTA SEGUNDO ORIGINÁRIO
        try:
            processo_originario_2 = navegador.find_element(
                by=By.XPATH,
                value='//*[@id="tableRelacionado"]/tbody/tr[2]/td[1]').text

            if len(processo_originario_2) >= 20:
                processo_originario_2 = processo_originario_2[0:-3]

            status2 = navegador.find_element(
                by=By.XPATH,
                value='//*[@id="tableRelacionado"]/tbody/tr[2]/td[3]').text

        except NoSuchElementException:
            processo_originario_2 = 'Sem dados de processo originário 2'
            status2 = 'Nulo'

        # TENTA TERCEIRO ORIGINÁRIO
        try:
            processo_originario_3 = navegador.find_element(
                by=By.XPATH,
                value='//*[@id="tableRelacionado"]/tbody/tr[3]/td[1]').text

            if len(processo_originario_3) >= 20:
                processo_originario_3 = processo_originario_3[0:-3]
            status3 = navegador.find_element(
                by=By.XPATH,
                value='//*[@id="tableRelacionado"]/tbody/tr[3]/td[3]').text

        except NoSuchElementException:
            processo_originario_3 = 'Sem dados de processo originário 3'
            status3 = 'Nulo'

        # COLETA DATA DE DISTRIBUIÇÃO
        sleep(0.6)
        data_distribuicao = navegador.find_element(by=By.ID, value='txtAutuacao').text  # COLETA DATA DE DISTRIBUIÇÃO
        sleep(0.5)
        data_reduzida = data_distribuicao[0:10]

        # ARMAZENAMENTO DE VALORES
        cj_processos.append(processo)
        cj_originarios.append(processo_originario_1)
        cj_precatorios.append(processo_originario_2)
        cj_originarios3.append(processo_originario_3)
        cj_datas.append(data_reduzida)
        cj_status1.append(status1)
        cj_status2.append(status2)
        cj_status3.append(status3)

        print(f'Contagem: {item}/{len(consultas)-1}')
        item += 1

    dicionario = {'Processo': cj_processos,
                  'originario_1': cj_originarios,
                  'Status 1': cj_status1,
                  'originario_2': cj_precatorios,
                  'Status 2': cj_status2,
                  'originario_3': cj_originarios3,
                  'Status 3': cj_status3,
                  'Data distribuição': cj_datas
                  }  # ARMAZENA OS VALORES EM UMA MATRIZ

    df = pd.DataFrame(dicionario)  # CRIA O DATAFRAME
    df.to_excel('./Lista de notas.xlsx')  # CRIA O ARQUIVO EXCEL
    return df


def consulta_base_de_dados(df_consulta):
    consulta_originario_no_principal = []
    for proc_format in df_consulta['Processo']:
        indice_principal = df_consulta.index[df_consulta['Processo'] == proc_format].tolist()
        originario = df_consulta.loc[indice_principal[0], 'originario_1']
        indice_originario = df_consulta.index[df_consulta['originario_1'] == originario].tolist()

        if originario in list(desdobramentos['numero']) and proc_format not in list(desdobramentos['numero']):
            indice_originario_principal = desdobramentos.index[desdobramentos['numero'] == originario].tolist()
            consulta_originario_no_principal.append((df_consulta['Processo'][indice_originario[0]],
                                                     df_consulta['originario_1'][indice_originario[0]],
                                                     desdobramentos.loc[indice_originario_principal[
                                                                            0], 'pasta_desdobramento'],
                                                     'originario_principal'))
    resultado_consulta = pd.DataFrame(consulta_originario_no_principal)
    resultado_consulta.to_excel('Resultado consulta.xlsx')


# EXECUÇÃO
options = webdriver.ChromeOptions()
options.add_experimental_option('prefs', {
    'download.default_directory': str(Path.cwd()),
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'plugins.always_open_pdf_externally': True,
})

navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)
navegador.implicitly_wait(5)
actions = ActionChains(navegador)


acessa_eproc()
download_planilha_comum()
cadastrados, consultas, desdobramentos = realiza_triagem()
df_notas = coleta_originarios(cadastrados, consultas, desdobramentos)
consulta_base_de_dados(df_notas)
navegador.quit()

# cria o modelo para SmartImport
prepara_import()
# re.enviar_email()

