# Importar a base de dados
import pandas as pd
import win32com.client as win32

# Visualizar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Faturamento por loja
tabela_faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

# Quantidade de produtos vendidos por loja
tabela_quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

# Calcular o ticket médio por loja
# Agrupar por ID Loja e calcular a média do Valor Final e da Quantidade
media_por_loja = tabela_vendas.groupby('ID Loja').agg({'Valor Final': 'mean', 'Quantidade': 'mean'})
# Calcular o ticket médio
media_por_loja['Ticket Médio'] = media_por_loja['Valor Final'] / media_por_loja['Quantidade']

# Criar tabelas em HTML usando pandas
faturamento_html = tabela_faturamento.to_html(index=True, header=True, float_format='R$ {:.2f}'.format)
quantidade_html = tabela_quantidade.to_html(index=True, header=True)

# Usar apenas a coluna 'Ticket Médio' para a tabela final
ticket_html = media_por_loja[['Ticket Médio']].to_html(index=True, header=True, float_format='R$ {:.2f}'.format)

# Enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'bluecrushangel@outlook.com'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = '''<h2> Relatório de vendas por loja </h2>''' + faturamento_html + '''<br>''' + quantidade_html + '''<br>''' + ticket_html
mail.Send()