def enviar_email():
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    import datetime
    import json

    data_atual = datetime.date.today().strftime("%d%m%y")

    # cria servidor
    host = 'smtp.gmail.com'
    port = 587
    # credenciais para login
    with open('credentials.json', 'r') as read_file:
        credenciais = json.load(read_file)

    username, password = credenciais['credentials']
    read_file.close()

    server = smtplib.SMTP(host, port)
    server.ehlo()
    server.starttls()
    server.login(username, password)

    # cria email
    corpo_email = 'Segue em anexo o arquivo para correção dos processos digitalizados no EPROC'
    msg = MIMEMultipart()
    msg['Subject'] = 'Correção digitalizados EPROC'
    msg['From'] = username
    msg['To'] = 'francis.calza@barbieriadvogados.com'
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
