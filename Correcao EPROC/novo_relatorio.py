def relatorio_email():

    import datetime
    import imaplib
    import email
    import json

    # data atual para filtrar emails
    data_atual = datetime.date.today().strftime("%d-%b-%Y")

    # credenciais para login
    with open('credentials.json', 'r') as read_file:
        credenciais = json.load(read_file)

    username, password = credenciais['credentials']

    # login no imap
    meu_email = imaplib.IMAP4_SSL('imap.gmail.com')
    meu_email.login(username, password)

    # Abre caixa de entrada
    meu_email.list()
    meu_email.select(mailbox='inbox', readonly=True)
    respostas, id_emails = meu_email.search(None, f'(ON "{data_atual}")')

    # Itera por cada email recebido
    for num in id_emails[0].split():
        # Coleta conteúdo do email em bytes
        resultado, dados = meu_email.fetch(num, '(RFC822)')
        texto = dados[0][1]
        texto = texto.decode('utf-8')
        texto_email = email.message_from_string(texto)
        # Percorre os bytes procurando anexo
        for parte in texto_email.walk():
            if parte.get_content_maintype() == 'multipart':
                continue
            if parte.get('Content-Disposition') is None:
                continue
            # Salva o anexo na pasta
            with open("relatorio_desdobramentos.xlsx", 'wb') as arquivo_excel:
                arquivo_excel.write(parte.get_payload(decode=True))
                arquivo_excel.close()
    print('Relatorio impresso')
