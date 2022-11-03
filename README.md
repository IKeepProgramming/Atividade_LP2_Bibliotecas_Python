# Biblioteca Pywin32 - Para automação.

  Essa biblioteca traz para o programador um pacote com diversas funções que conectam o programa Python com os aplicativos do computador, para isso é utilizado uma API do próprio Windows, que garante que essa conexão sejá feita, como exemplo, vou utilizar o Outlook para fazer o envio de um e-mail apenas rodando o programa.
  
# Envio de e-mail com o Pywin32:

import win32com.client as win32

# Criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

# Criar um email
email = outlook.CreateItem(0)

faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

# Configurar as informações do seu e-mail
email.To = "amaraz228@gmail.com; amaral.carlos@ifsp.edu.br"
email.Subject = "Automação de E-mails com o Pywin32"
email.HTMLBody = f"""
<p>Fala rapaziada, boa noite, e-mail recebido com Sucesso</p>

<p>Nossa loja faturou um total de R${faturamento}</p>
<p>Vendemos {qtde_produtos} produtos</p>
<p>A média em dinheiro foi de R${ticket_medio}</p>

<p>Abs,</p>
<p>Obs: Funciona pra estatisticas de rede também</p>
"""

anexo = "COLOCAR AQUI O CAMINHO DO ANEXO"
email.Attachments.Add(anexo)

email.Send()
print("Email Enviado")
