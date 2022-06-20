# Importar a base de dados
import pandas as pd
import win32com.client as win32

# Visualizar a base de dados
tab_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_columns', None) #=> Mostra todas as colunas no terminal

# Faturamento por loja
faturamento = tab_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

# Quantidade de produtos vendidos por loja
qtd_produtos_vendidos = tab_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtd_produtos_vendidos)
print('-' * 50)

# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / qtd_produtos_vendidos['Quantidade']).to_frame()
print(ticket_medio)

# Enviar email com relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'gabrielmedsilva@outlook.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = '''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{qtd_produtos_vendidos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att..</p>
<p>Gabriel</p>
'''

mail.Send()