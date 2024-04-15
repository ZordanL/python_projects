<h1><b> Passo a passo para resolver o problema </h1></b>
Passo 0: Entender o problema e suas restrições<br/>
Passo 1: Importar a base de dados<br/>
Passo 2: Calcular o produto mais vendido (quantidade) <br/>
Passo 3: Calcular o produto mais vendido (faturamento em reais) <br/>
Passo 4: Calcular o loja que mais vendeu (faturamento em reais) <br/>
Passo 5: Calcular o ticket médio por loja <br/>
Passo 6: Criar um dasboard da loja que mais vendeu (faturamento reais) <br/>
Passo 7: Enviar um email automátiico <br/>

<i> Blibiotecas exploradas: Pandas, Plotly, Smtplib, Babel.Numbers </i>

#CRIAÇÃO DOS INDICADORES
"""

#Importação das bibliotecas

import pandas as pd
import plotly.express as px
import smtplib
import email.message as em

#Passo 1: Importar a base de dados

tabela_vendas=pd.read_excel("/content/drive/MyDrive/Udemy_PythonDS/Project 1/Vendas.xlsx")
tabela_vendas = tabela_vendas.rename(columns={'Valor Unitário':'Valor_Unitario','Cód.':'Cod'})

# Passo 2: Calcular o produto mais vendido (quantidade)

tb_quantidade_produto= tabela_vendas.groupby('Produto').sum() #por conta da função groupby a coluna que faz o agrupamento vira index- indexador unico
tb_quantidade_produto= tb_quantidade_produto[['Quantidade']]
tb_quantidade_produto = tb_quantidade_produto.sort_values(by='Quantidade', ascending=False)

#Passo 3: Calcular o produto mais vendido (faturamento em reais)

tabela_vendas['Faturamento_Reais'] = tabela_vendas['Quantidade']*tabela_vendas['Valor_Unitario']
tb_faturamento_produto = tabela_vendas.groupby('Produto').sum()
tb_faturamento_produto = tb_faturamento_produto[['Faturamento_Reais']]
tb_faturamento_produto = tb_faturamento_produto.sort_values(by='Faturamento_Reais', ascending=False)

# Passo 4: Calcular o loja que mais vendeu (faturamento em reais)
tb_faturamento_loja = tabela_vendas.groupby('Loja').sum()
tb_faturamento_loja = tb_faturamento_loja[['Faturamento_Reais']]
tb_faturamento_loja = tb_faturamento_loja.sort_values(by='Faturamento_Reais', ascending=False)

# Passo 5: Calcular o ticket médio por loja]

tabela_vendas['Ticket_Medio'] = tabela_vendas['Valor_Unitario']
tb_ticket_loja = tabela_vendas.groupby('Loja').mean(numeric_only=True)
tb_ticket_loja = tb_ticket_loja[['Ticket_Medio']]
tb_ticket_loja = tb_ticket_loja.sort_values(by='Ticket_Medio', ascending=False)

# Passo 6: Criar um dasboard da loja que mais vendeu (faturamento reais)

grafico = px.bar(tb_faturamento_loja, y='Faturamento_Reais', x=tb_faturamento_loja.index) #por conta da função groupby a coluna loja virou index

"""#FORMATAÇÃO NUMÉRICA E MONETÁRIA"""

from babel.numbers import format_currency
#As formatações devem ser feitas separadamente e após os cálculos, em células separadas, porque fazem com que os valores passem a ser texto

#Faturamento por loja
tb_faturamento_loja_formatado = pd.DataFrame(tb_faturamento_loja['Faturamento_Reais'].apply(lambda x: format_currency(x,'BRL',locale='pt_BR'))) #formatar moeda
tb_faturamento_loja_formatado = tb_faturamento_loja_formatado.reset_index() #organizar cabeçalho, resetando index
tb_faturamento_loja_formatado_html = tb_faturamento_loja_formatado.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">') #se torna uma estrutura html e não df pandas

#Quantidade vendida por produto
tb_quantidade_produto_formatado = pd.DataFrame(tb_quantidade_produto['Quantidade'].apply(lambda x: '{:,.2f}'.format(x).replace(',','.'))) #formatar numero
tb_quantidade_produto_formatado = tb_quantidade_produto_formatado.reset_index()
tb_quantidade_produto_formatado_html = tb_quantidade_produto_formatado.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">')

#Faturamento por produto
tb_faturamento_produto_formatado = pd.DataFrame(tb_faturamento_produto['Faturamento_Reais'].apply(lambda x: format_currency(x,'BRL',locale='pt_BR')))
tb_faturamento_produto_formatado = tb_faturamento_produto_formatado.reset_index()
tb_faturamento_produto_formatado_html = tb_faturamento_produto_formatado.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">')

#Ticket médio por loja
tb_ticket_loja_formatado = pd.DataFrame(tb_ticket_loja['Ticket_Medio'].apply(lambda x: format_currency(x,'BRL',locale='pt_BR')))
tb_ticket_loja_formatado = tb_ticket_loja_formatado.reset_index()
tb_ticket_loja_formatado_html = tb_ticket_loja_formatado.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">')

"""#ENVIAR EMAIL"""

# Passo 7: Enviar um email automático
#Utilizando 3 aspas para guardar várias linhas dentro de uma varável
#Utilizar o f para acrescentar variáveis no corpo do email


corpo_email = f"""
<p>E aí pessoal, tudo bem?</p>

<p>Seguem indicadores de venda</p>

<p><b>Faturamento por loja</p></b>
<p>{tb_faturamento_loja_formatado_html}</p>

<p><b>Quantidade vendida por produto</b></p>
<p>{tb_quantidade_produto_formatado_html}</p>

<p><b>Faturamento por produto</p>
<p>{tb_faturamento_produto_formatado_html} </b></p>

<p><b>Ticket médio por loja</p></b>
<p>{tb_ticket_loja_formatado_html} </p>

<p>Qualquer duvida estou à disposição</p>
<p>Abs.,</p>
<p>Lucas Zordan</p>
"""

#Configurações para envio
msg = em.Message()
msg['Subject'] = "Relatório de vendas" #Assunto do email
msg['From'] = 'your_email@gmail.com' #Email que vai disparar
msg['To'] = 'yout_email@gmail.com'#Email que vai receber
password = 'your_password' #Senha do email que vai enviar
msg.add_header('Content-Type', 'text/html') #Formato do conteúdo
msg.set_payload(corpo_email) #Mensagem

#Configurações do servidor - GMAIL
s = smtplib.SMTP('smtp.gmail.com: 587') #Simple Mail Transfer Protocol
s.starttls() #criotografia

#Credenciais do login
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
print('Email enviado')
