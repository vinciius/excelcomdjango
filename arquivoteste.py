# Importar a base de dados
import pandas as pd
from prettytable import PrettyTable

# Visualizar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Faturamento por loja
tabela_faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# Quantidade de produtos vendidos por loja
tabela_quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Ticket médio por produto em cada loja
tabela_ticket_medio = tabela_vendas['Valor Final'] / tabela_vendas['Quantidade']

# Enviar um email com o relatório


# Tabela personalizada de ticket médio
ticket_table = PrettyTable()
ticket_table.field_names = ["ID Loja", "Ticket Médio"]
for index, row in tabela_ticket_medio.items():
    ticket_table.add_row([index, f"R$ {row:.2f}"])

# Tabela personalizada de faturamento
faturamento_table = PrettyTable()
faturamento_table.field_names = ["ID Loja", "Valor Final"]
for index, row in tabela_faturamento.iterrows():
    faturamento_table.add_row([row.name, f"R$ {row['Valor Final']:.2f}"])

# Tabela personalizada de quantidade
quantidade_table = PrettyTable()
quantidade_table.field_names = ["ID Loja", "Quantidade Vendida"]
for index, row in tabela_quantidade.iterrows():
    quantidade_table.add_row([row.name, row['Quantidade']])

# Imprimir as tabelas
print("\n" + "="*40)
print("Faturamento por Loja")
print("="*40)
print(faturamento_table)

print("\n" + "="*40)
print("Quantidade de Produtos Vendidos por Loja")
print("="*40)
print(quantidade_table)
print("="*40)

print("\n" + "="*40)
print("Ticket Médio por Loja")
print("="*40)
print(ticket_table)