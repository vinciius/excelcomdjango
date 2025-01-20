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
faturamento_html = tabela_faturamento.reset_index().to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format}, index=False, header=True)
quantidade_html = tabela_quantidade.reset_index().to_html(index=False, header=True)
ticket_html = media_por_loja[['Ticket Médio']].reset_index().to_html(formatters={'Ticket Médio': 'R$ {:,.2f}'.format}, index=False, header=True)

# Adicionar estilo CSS para centralizar o texto e os cabeçalhos
faturamento_html = '<style>table {text-align: center;} th {text-align: center;}</style>' + faturamento_html
quantidade_html = '<style>table {text-align: center;} th {text-align: center;}</style>' + quantidade_html
ticket_html = '<style>table {text-align: center;} th {text-align: center;}</style>' + ticket_html

# Enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'bluecrushangel@outlook.com'
mail.Subject = 'Relatório de vendas por loja'
mail.HTMLBody = '''<h2> Relatório de vendas por loja </h2>''' + '''<p> <b> Faturamento por loja: </b> </p>''' + faturamento_html + '''<br>''' + '''<p> <b> Quantidade de produtos vendidos por loja: </b> </p>''' + quantidade_html + '''<br>''' + '''<p> <b> Ticket Médio por loja: </b> </p>''' + ticket_html 
mail.Send()