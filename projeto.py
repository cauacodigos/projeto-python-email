import pandas as pd 
import win32com.client as win32


tabela_vendas = pd.read_excel('Vendas.xlsx') 


pd.set_option('display.max_columns' , None) 

print(tabela_vendas)


faturamento = tabela_vendas[['ID Loja' , 'Valor Final']].groupby('ID Loja').sum() 

print(faturamento)


quantidade = tabela_vendas [['ID Loja' , 'Quantidade']].groupby('ID Loja').sum()

print(quantidade)

print ('-' *50)

ticket_medio =  (faturamento['Valor Final']  / quantidade['Quantidade']).to_frame() 

ticket_medio = ticket_medio.rename(columns={0: 'Ticket medio'})

print(ticket_medio)



outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.to = 'abcdefg@gmail.com'
mail.Subject = 'Relatorio de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>



<p>Segue o relatorio de Vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade vendida:</p>  
{quantidade.to_html()}

<p>Ticket médio por produto:</p>
{ticket_medio.to_html(formatters={'Ticket médio': 'R${:,.2f}'.format})}

<p>att.,</p> 
<p>Cauã<p>
'''

mail.Send()

print('Email enviado')