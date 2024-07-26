import win32com.client as win32
from emails import e_mail, email2

outlook = win32.Dispatch('outlook.application')

email = outlook.CreateItem(0)

vendas_totais = 3000
quant_produtos = 25
vendas_por_produto = vendas_totais/quant_produtos

#Informando em qual e_mail será o destinatário
email.To = e_mail; email2

#Informa o assunto do e_mail
email.Subject = 'Teste de e_mail automático pelo python'

#Informa o conteúdo do e_e_mail sendo feito por HTML 
email.HTMLBody = f'''
    <p> Estou enviando uma e-mail teste automatico para testar meus onhecimentos em python o valor total de vendas por produto ficou em {vendas_por_produto}R$'''

#envia o e_mail
email.Send()

print('Parabens deu certo')