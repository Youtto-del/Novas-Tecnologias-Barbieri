def relatorio_email():

    import datetime
    from imap_tools import MailBox, AND

    # READ ME - https://github.com/ikvk/imap_tools#id6

    # data atual para filtrar emails
    data_atual = datetime.date.today().strftime("%d%m%y")

    # credenciais para login
    username = "navarroreports@gmail.com"
    password = "dezwjuzigsmchbrr"

    # lista de imaps: https://www.systoolsgroup.com/imap/
    # login no imap
    meu_email = MailBox('imap.gmail.com').login(username, password)

    # pegar emails com um anexo específico
    lista_emails = meu_email.fetch(AND(from_="barbieri@barbieriadvogados.com"))
    for email in lista_emails:
        if len(email.attachments) > 0 and email.date.strftime('%d%m%y') == data_atual:
            print('Assunto:', email.subject)
            print('Corpo:', email.text)
            print('Data:', email.date.strftime('%d%m%y'))
            for anexo in email.attachments:
                if "Smart Report" in anexo.filename:
                    print('Tipo de conteúdo:', anexo.content_type)
                    with open("relatorio_desdobramentos.xlsx", 'wb') as arquivo_excel:
                        arquivo_excel.write(anexo.payload)
    return


def enviar_email():
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    import datetime

    data_atual = datetime.date.today().strftime("%d%m%y")

    # cria servidor
    host = 'smtp.gmail.com'
    port = '587'
    login = 'navarroreports@gmail.com'
    senha = 'SenhaReport&2023'

    server = smtplib.SMTP(host, port)
    server.ehlo()
    server.starttls()
    server.login(login, senha)

    # cria email
    corpo_email = 'Segue em anexo o arquivo para correção dos processos digitalizados no EPROC'
    msg = MIMEMultipart()
    msg['Subject'] = 'Correção digitalizados EPROC'
    msg['From'] = login
    msg['To'] = 'felipensamaral@gmail.com' # francis.calza@barbieriadvogados.com
    msg.attach(MIMEText(corpo_email, 'Plain'))

    # adiciona anexos
    local_anexo = rf'.\SmartImports\Correcao Digit EPROC ATT - {data_atual}.xlsx'
    anexo = open(local_anexo, 'rb')

    att = MIMEBase('application', 'octet-stream')
    att.set_payload(anexo.read())
    encoders.encode_base64(att)

    att.add_header('Content-Disposition', f'attachment; filename=Correcao Digit EPROC ATT - {data_atual}.xlsx')
    anexo.close()

    msg.attach(att)

    # enviar email no servidor SMTP
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    server.quit()

    print('Email enviado')

    return


enviar_email()
