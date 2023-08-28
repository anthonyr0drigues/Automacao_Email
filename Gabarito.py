import pandas as pd
import win32com.client as win32
from Tools.scripts.dutree import display

# importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')


# visualizar base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-' * 50)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

# quantidade vendido por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)

# ticket médio por loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0:'Ticket Médio'})
print(ticket_medio)
print("-" * 50)

# enviar email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.to = 'anthonyrodriguesoficial@yahoo.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.htmlBody = f'''
<p>Prezados,</p>

<p>Segue o relatório de vendas por cada loja.<p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou á disposição.</p>

Att;
Anthony
'''
mail.Send()
print('Email Enviado')
