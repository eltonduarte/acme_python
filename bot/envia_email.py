import win32com.client as win32

def email_conclusao(destino):

    # criar a integração com o outlook
    outlook = win32.Dispatch('outlook.application')

    # criar um email
    email = outlook.CreateItem(0)

    faturamento = 1500
    qtde_produtos = 10
    ticket_medio = faturamento / qtde_produtos

    # configurar as informações do seu e-mail
    email.To = destino
    email.Subject = "E-mail automático do Python"
    email.HTMLBody = f"""
    <p>Prezado(as)</p>

    <p>Segue relatório em anexo</p>


    <p>Atenciosamente</p>
    <p>Bot</p>
    """

    anexo = r"C:\acme\Relatorio.xlsx"
    email.Attachments.Add(anexo)

    email.Send()
    print("Email Enviado")