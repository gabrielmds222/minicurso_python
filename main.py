# Importar a base de dados
import pandas as pd
tab_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
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