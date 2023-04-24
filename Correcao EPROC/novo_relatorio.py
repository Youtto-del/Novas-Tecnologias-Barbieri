import imaplib


def relatorio_email():

    import datetime
    import imaplib
    import email

    # data atual para filtrar emails
    data_atual = datetime.date.today().strftime("%d-%b-%Y")

    # credenciais para login
    username = "relatoriosnavarro@gmail.com"
    password = "vfgmigllhkuevzdn"

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


# def enviar_email():
#     import smtplib
#     from email.mime.multipart import MIMEMultipart
#     from email.mime.text import MIMEText
#     from email.mime.base import MIMEBase
#     from email import encoders
#     import datetime
#
#     data_atual = datetime.date.today().strftime("%d%m%y")
#
#     # cria servidor
#     host = 'smtp.gmail.com'
#     port = '587'
#     login = 'navarroreports@gmail.com'
#     senha = 'SenhaReport&2023'
#
#     server = smtplib.SMTP(host, port)
#     server.ehlo()
#     server.starttls()
#     server.login(login, senha)
#
#     # cria email
#     corpo_email = 'Segue em anexo o arquivo para correção dos processos digitalizados no EPROC'
#     msg = MIMEMultipart()
#     msg['Subject'] = 'Correção digitalizados EPROC'
#     msg['From'] = login
#     msg['To'] = 'felipensamaral@gmail.com' # francis.calza@barbieriadvogados.com
#     msg.attach(MIMEText(corpo_email, 'Plain'))
#
#     # adiciona anexos
#     local_anexo = rf'.\SmartImports\Correcao Digit EPROC ATT - {data_atual}.xlsx'
#     anexo = open(local_anexo, 'rb')
#
#     att = MIMEBase('application', 'octet-stream')
#     att.set_payload(anexo.read())
#     encoders.encode_base64(att)
#
#     att.add_header('Content-Disposition', f'attachment; filename=Correcao Digit EPROC ATT - {data_atual}.xlsx')
#     anexo.close()
#
#     msg.attach(att)
#
#     # enviar email no servidor SMTP
#     server.sendmail(msg['From'], msg['To'], msg.as_string())
#     server.quit()
#
#     print('Email enviado')
#
#     return
#
relatorio_email()
# enviar_email()
