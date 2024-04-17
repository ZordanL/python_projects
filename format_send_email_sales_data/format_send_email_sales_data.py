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

#step 4: Calculate the store that sold the most (profit)

df_profit_store = df_sales.groupby('Store').sum()
df_profit_store = df_profit_store[['Profit']]
df_profit_store = df_profit_store.sort_values(by='Profit', ascending=False)

#step 5: Calculate the average ticket per store

df_sales['Average_ticket'] = df_sales['Unitary_value']
df_ticket_store = df_sales.groupby('Store').mean(numeric_only=True)
df_ticket_store = df_ticket_store[['Average_ticket']]
df_ticket_store = df_ticket_store.sort_values(by='Average_ticket', ascending=False)

#step 6: Create a dashboard of the store that sold the most (profit)

graph = px.bar(df_profit_store, y='Profit', x=df_profit_store.index) #Due to the groupby function, the store column became index

#NUMERICAL AND MONETARY FORMATTING

from babel.numbers import format_currency
#Formatting must be done separately and after calculations, in separated cells, because they make the values become text

#Porfit per Store
df_profit_store_formatted = pd.DataFrame(df_profit_store['Profit'].apply(lambda x: format_currency(x,'BRL',locale='pt_BR'))) #format currency
df_profit_store_formatted = df_profit_store_formatted.reset_index() #organizing header by resetting index
df_profit_store_formatted_html = df_profit_store_formatted.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">') #it becomes an html structure and not pandas df

#Quantity sold per product
df_quantity_product_formatted = pd.DataFrame(df_quantity_product['Quantity'].apply(lambda x: '{:,.2f}'.format(x).replace(',','.'))) #format number
df_quantity_product_formatted = df_quantity_product_formatted.reset_index()
df_quantity_product_formatted_html = df_quantity_product_formatted.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">')

#Profit per product
df_profit_product_formatted = pd.DataFrame(df_profit_product['Profit'].apply(lambda x: format_currency(x,'BRL',locale='pt_BR')))
df_profit_product_formatted = df_profit_product_formatted.reset_index()
df_profit_product_formatted_html = df_profit_product_formatted.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">')

#Average Ticket per Store
df_ticket_store_formatted = pd.DataFrame(df_ticket_store['Average_ticket'].apply(lambda x: format_currency(x,'BRL',locale='pt_BR')))
df_ticket_store_formatted = df_ticket_store_formatted.reset_index()
df_ticket_store_formatted_html = df_ticket_store_formatted.to_html(index=False,justify='center',border='0').replace('<tbody>', '<tbody style= "text-align: center; color: #484848; background: #F7F7F7">')

#SEND EMAIL

# Step 7: Send an automatic email
#Using 3 quotes to save multiple lines within a variable
#Use f to add variables to the body of the email

email_body = f"""
<p>Hi, there! I hope this message finds you well</p>

<p>The sales indicator follows int the body of the email</p>

<p><b>Profit per Store</p></b>
<p>{df_profit_store_formatted_html}</p>

<p><b>Quantity sold per product</b></p>
<p>{df_quantity_product_formatted_html}</p>

<p><b>Profit per product</p>
<p>{df_profit_product_formatted_html} </b></p>

<p><b>Average Ticket per Store</p></b>
<p>{df_ticket_store_formatted_html} </p>

<p>"Any questions, I'm at your disposal. Regards,"</p>
<p>Signature</p>
"""

#Settings for sending the email
msg = em.Message()
msg['Subject'] = "Sales Report" #Subject
msg['From'] = 'from_email@gmail.com'
msg['To'] = 'to_email@gmail.com'
password = 'your_password' #Password for the account that will send the email
msg.add_header('Content-Type', 'text/html') #Content format
msg.set_payload(email_body) #Message

#Server Settings - GMAIL
s = smtplib.SMTP('smtp.gmail.com: 587') #Simple Mail Transfer Protocol
s.starttls() #cryotography

#Login credentials
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
print('Email sent')
