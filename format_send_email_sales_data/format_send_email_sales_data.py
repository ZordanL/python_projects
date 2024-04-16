#importing libraries

import pandas as pd
import plotly.express as px
import smtplib
import email.message as em

#step 1: import database

df_sales=pd.read_excel("/content/drive/MyDrive/Udemy_PythonDS/Project 1/Sales.xlsx")

#step 2: Calculate the best-selling product (quantity)

df_quantity_product= df_sales.groupby('Product').sum() #Due to the groupby function, the column that does the grouping becomes an index - unique indexer
df_quantity_product= df_quantity_product[['Quantity']]
df_quantity_product = df_quantity_product.sort_values(by='Quantity', ascending=False)

#step 3: Calculate the best-selling product (profit)

df_sales['Profit'] = df_sales['Quantity']*df_sales['Unitary_value']
df_profit_product = df_sales.groupby('Product').sum()
df_profit_product = df_profit_product[['Profit']]
df_profit_product = df_profit_product.sort_values(by='Profit', ascending=False)

# Passo 4: Calculate the store that sold the most (profit)

df_profit_store = df_sales.groupby('Store').sum()
df_profit_store = df_profit_store[['Profit']]
df_profit_store = df_profit_store.sort_values(by='Profit', ascending=False)

# Passo 5: Calculate the average ticket per store

df_sales['Average_ticket'] = df_sales['Unitary_value']
df_ticket_store = df_sales.groupby('Store').mean(numeric_only=True)
df_ticket_store = df_ticket_store[['Average_ticket']]
df_ticket_store = df_ticket_store.sort_values(by='Average_ticket', ascending=False)

# Passo 6: Create a dashboard of the store that sold the most (profit)

graph = px.bar(df_profit_store, y='Profit', x=df_profit_store.index) #Due to the groupby function, the store column became index

#NUMERICAL AND MONETARY FORMATTING

from babel.numbers import format_currency
#Formatting must be done separately and after calculations, in separated cells, because they make the values become text

#Faturamento por Store
df_profit_store_formatted = pd.DataFrame(df_profit_store['Profit'].apply(lambda x: format_currency(x,'BRL',locale='pt_BR'))) #formatar moeda
df_profit_store_formatted = df_profit_store_formatted.reset_index() #organizar cabeçalho, resetando index
df_profit_store_formatted_html = df_profit_store_formatted.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">') #se torna uma estrutura html e não df pandas

#Quantidade vendida por produto
df_quantity_product_formatted = pd.DataFrame(df_quantity_product['Quantity'].apply(lambda x: '{:,.2f}'.format(x).replace(',','.'))) #formatar numero
df_quantity_product_formatted = df_quantity_product_formatted.reset_index()
df_quantity_product_formatted_html = df_quantity_product_formatted.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">')

#Faturamento por produto
df_profit_product_formatted = pd.DataFrame(df_profit_product['Profit'].apply(lambda x: format_currency(x,'BRL',locale='pt_BR')))
df_profit_product_formatted = df_profit_product_formatted.reset_index()
df_profit_product_formatted_html = df_profit_product_formatted.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">')

#Ticket médio por Store
df_ticket_store_formatted = pd.DataFrame(df_ticket_store['Average_ticket'].apply(lambda x: format_currency(x,'BRL',locale='pt_BR')))
df_ticket_store_formatted = df_ticket_store_formatted.reset_index()
df_ticket_store_formatted_html = df_ticket_store_formatted.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">')

"""#ENVIAR EMAIL"""

# Passo 7: Enviar um email automático
#Utilizando 3 aspas para guardar várias linhas dentro de uma varável
#Utilizar o f para acrescentar variáveis no corpo do email


corpo_email = f"""
<p>E aí pessoal, tudo bem?</p>

<p>Seguem indicadores de venda</p>

<p><b>Faturamento por Store</p></b>
<p>{df_profit_store_formatted_html}</p>

<p><b>Quantidade vendida por produto</b></p>
<p>{df_quantity_product_formatted_html}</p>

<p><b>Faturamento por produto</p>
<p>{df_profit_product_formatted_html} </b></p>

<p><b>Ticket médio por Store</p></b>
<p>{df_ticket_store_formatted_html} </p>

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
