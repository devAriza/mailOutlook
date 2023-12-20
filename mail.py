import win32com.client as win32

def enviar_correo_outlook(asunto, cuerpo, destinatario):
    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.Subject = asunto
    mail.Body = cuerpo
    mail.To = destinatario

    
    mail.Attachments.Add(r'ruta de archivo adjunto deseado')
    mail.SentOnBehalfOfName = cuenta_origen
    mail.Send()


asunto = '#asunto#'
cuerpo = 'Escribir la redacción del correo electrónico'
destinatario = 'email destinatario'
cuenta_origen = 'email de computadora desde Outlook'

enviar_correo_outlook(asunto, cuerpo, destinatario)
