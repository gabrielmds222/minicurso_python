# Importar a base de dados
import pandas as pd
tab_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
pd.set_option('display.max_columns', None) #=> Mostra todas as colunas no terminal

# Faturamento por loja
faturamento = tab_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
# Quantidade de produtos vendidos por loja

# Ticket médio por produto em cada loja

# Enviar email com relatório