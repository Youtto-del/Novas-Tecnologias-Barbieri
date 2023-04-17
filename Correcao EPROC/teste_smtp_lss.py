def relatorio_email():

    import datetime
    from imap_tools import MailBox, AND

    # READ ME - https://github.com/ikvk/imap_tools#id6

    # data atual para filtrar emails
    data_atual = datetime.date.today().strftime("%d%m%y")

    # credenciais para login
    username = "navarroreports@gmail.com"
    password = "SenhaReport&2023"

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


# def enviar_email():
#     import smtplib
#     from email.message import EmailMessage
#     import datetime
#
#     data_atual = datetime.date.today().strftime("%d%m%y")
#
#     msg = EmailMessage()
#     msg['Subject'] = 'Correção digitalizados EPROC'
#     msg['From'] = 'navarroreports@gmail.com'
#     msg['To'] = 'felipensamaral@gmail.com' #francis.calza@barbieriadvogados.com
#     msg.set_content('Segue em anexo o arquivo para correção dos processos digitalizados no EPROC')
#
#     with open(rf'.\SmartImports\Correcao Digit EPROC ATT - {data_atual}.xlsx', 'rb') as conteudo_arquivo:
#         anexo = conteudo_arquivo.read()
#         msg.add_attachment(anexo,
#                            maintype='application',
#                            subtype='xlsx',
#                            filename=f'Correcao Digit EPROC ATT - {data_atual}.xlsx')
#
#     senha = 'dezwjuzigsmchbrr'
#
#     with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
#         smtp.login(msg['From'], senha)
#         smtp.send_message(msg)
#
#     print('Email enviado')
#
#     return
#
#
# enviar_email()
