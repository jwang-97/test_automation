import win32com.client as win32

outlook = win32.Dispatch('outlook.application')

mail = outlook.CreateItem(0)
mail.To = 'JWang294@slb.com'
mail.Subject = 'Sample Email'
mail.HTMLBody = '<h3>This is HTML Body</h3>'
mail.Body = "This is the normal body"
mail.Attachments.Add(r'C:\Users\JWang294\Documents\projectfile\config.json')
# mail.Attachments.Add('c:\\sample2.xlsx')
mail.CC = 'JWang294@slb.com'