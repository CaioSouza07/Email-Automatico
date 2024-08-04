import win32com.client as win32
import globals



def enviarEmail():

    
    # criar a integração com o outlook
    outlook = win32.Dispatch('Outlook.Application')

    # criar um email
    email = outlook.CreateItem(0)

    # configurar as informações do seu e-mail
    email.To = "caiodesouza.cds@gmail.com"
    email.Subject = "Planilha de Produtos Zerados"

    imgAssinatura = "C://Users/DESKTOP/Desktop/Email-Automatico/img/assinaturateste.jpg"
    assinatura = email.Attachments.Add(imgAssinatura)
    assinatura.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "myimage")

    email.HTMLBody = f"""
    <p>Boa tarde,</p>

    <p>Segue em anexo a planilha com os produtos zerados do dia {globals.dataAtual}.</p>

    <p>Permaneço a disposição,</p>

    <p>Atenciosamente.</p>

    <img src="cid:myimage">
    """

    anexo = "C://Users/DESKTOP/arquivosCaio/automatico.xlsx"
    email.Attachments.Add(anexo)

    email.Send()
    print("Email Enviado")
    
    globals.loop = False

    