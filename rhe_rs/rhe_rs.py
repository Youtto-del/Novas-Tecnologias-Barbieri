from time import sleep
import pandas as pd
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from pathlib import Path
import tempfile
import shutil

# CONFIGURA
options = webdriver.ChromeOptions()
download_dir = tempfile.mkdtemp()

options.add_experimental_option('prefs', {
    'download.default_directory': download_dir,
    'download.prompt_for_download': False,
    'download.directory_upgrade': True,
    'plugins.always_open_pdf_externally': True,
})


navegador = webdriver.Chrome(ChromeDriverManager().install(), options=options)
navegador.implicitly_wait(5)
actions = ActionChains(navegador)

# INPUT
dados = pd.read_excel('input.xlsx')
print(dados)

for idx, row in dados.iterrows():
    id_cliente = row['id']
    senha = row['senha']
    data = row['termo_inicial']
    print(id_cliente, senha, data)
    p = Path(f'./Resultados/{id_cliente}')
    p.mkdir()

    # NAVEGADOR
    navegador.get('https://secweb.procergs.com.br/rheportal/logon.xhtml')
    navegador.find_element(By.ID, value='formLogin_:matriculaWeb').send_keys(id_cliente)
    navegador.find_element(By.ID, value='formLogin_:senhaWeb').send_keys(senha)
    navegador.find_element(By.ID, value='formLogin_:entrarWeb').click()
    sleep(1.5)
    navegador.get('https://secweb.procergs.com.br/rheportal/pages/contracheque/contracheque-list.xhtml')
    sleep(1.5)

    # CONSULTA DATA DO ULTIMO CONTRACHEQUE PARA DATA FINAL
    ultimo_contracheque = navegador.find_element(By.XPATH, value='//*[@id="form:lista_data"]/tr[1]/td[1]/a').text
    mes_final = ultimo_contracheque[:2]
    ano_final = ultimo_contracheque[-4:]
    print('Mês e ano finais:', mes_final, ano_final)

    # PERÍODO INICIAL
    ano_inicial = data[-4:]
    mes_inicial = data[:2]
    print('Mês e ano iniciais:', mes_inicial, ano_inicial)

    # FILTRO PARA 100 RESULTADOS
    select_element = navegador.find_element(By.XPATH, value='//*[@id="form:lista:j_id2"]')
    select = Select(select_element)
    select.select_by_value('100')
    sleep(2)

    # CONTABILIZA PÁGINAS
    paginas = navegador.find_element(By.CLASS_NAME, value='ui-paginator-current').text.split('/')[1].split(' ')[0]
    print('paginas', paginas)

    # CRIA BASE PARA DATASET DE ID
    lista_info_contracheques = []
    for pg in range(0, int(paginas)):
        trs = navegador.find_elements(By.XPATH, value='//*[@id="form:lista_data"]/tr')      # TRS PARA TEXTOS
        spans = navegador.find_elements(By.XPATH, value='//*[@id="form:lista_data"]/tr/td[1]/a')        # SPANS PARA HREFS
        print(len(trs))
        print(len(spans))

        cont = 1
        for tr in trs:
            data_tr = tr.text.split(' ')[0]
            tipo_tr = tr.text[7:]
            local_span = navegador.find_element(By.XPATH, value=f'//*[@id="form:lista_data"]/tr[{cont}]/td[1]/a')

            id_contracheque = local_span.get_attribute('href').split('id=')[1].split('&')[0]
            lista_info_contracheques.append((id_contracheque, data_tr, tipo_tr))
            print(id_contracheque, data_tr, tipo_tr)
            cont += 1

        try:
            navegador.find_element(By.XPATH, value='//*[@id="form:lista_paginator_bottom"]/a[3]').click()
        except:
            pass
        sleep(1)

    df = pd.DataFrame(lista_info_contracheques, columns=['id', 'data', 'tipo'])
    print(df)
    df.to_excel('Resultado.xlsx')
    print(len(lista_info_contracheques), 'Contracheques encontrados')


    # ITERA PELOS IDS ENCONTRADOS E BAIXA OS PDFS
    for i, r in df.iterrows():
        if r['data'] == data:
            break
        correcao_tipo = r['tipo'].replace(' ', '+')
        link_financeiro = f'https://secweb.procergs.com.br/rheportal/pages/contracheque/contracheque-form.xhtml?id={r["id"]}&num_folha=1&emp_codigo=1&mes_folha={r["data"][:2]}&ano_folha={r["data"][-4:]}&nome_folha={correcao_tipo}'
        print(link_financeiro)

        navegador.get(url=link_financeiro)
        navegador.find_element(By.ID, value='j_idt131_menuButton').click()
        sleep(0.5)
        navegador.find_element(By.ID, value='j_idt133').click()
        sleep(3)

        nome_arquivo = rf'.\Resultados\{id_cliente}\{r["data"][-4:]}-{r["data"][:2]}-{correcao_tipo}-{id_cliente}.pdf'

        try:
            while len(list(Path(download_dir).glob('*.pdf'))) == 0:
                sleep(1)  # espera o download terminar
            # pega o 1o pdf que tiver, só terá 1 pois a pasta estava vazia antes:
            arquivo = list(Path(download_dir).glob('*.pdf'))[0]
            shutil.move(arquivo, nome_arquivo)
        finally:
            shutil.rmtree(download_dir)  # remove todos os arquivos temporários

navegador.quit()
print('FINALIZADO')


