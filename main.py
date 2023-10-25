import win32com.client as win32

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)  # 0 representa um e-mail
mail.Display()
mail.Subject = 'Assunto do e-mail'
mail.Body = f'Corpo do e-mail'

# Adicione os destinatários (utilizar ; para multiplos e-mails)
mail.To = '...@....com.br'
print(outlook)


# Adicione os anexos
# mail.Attachments.Add(r'C:\Caminho\arquivo.txt')

# Envie o e-mail
mail.Send()
