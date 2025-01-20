# Projeto Django: Base de Aprendizado

## Introdução
Este projeto é uma aplicação Django desenvolvida como uma base de aprendizado para futuros projetos. Ele serve como um ponto de partida para desenvolvedores que desejam entender a estrutura e os componentes fundamentais de um projeto Django.

## Objetivo
O objetivo deste projeto é fornecer um exemplo prático de como configurar e estruturar uma aplicação Django. Ele inclui exemplos de como integrar um arquivo Excel para manipulação de dados, demonstrando a funcionalidade básica de importação e processamento de dados.

## Estrutura do Projeto
- **arquivoteste.py**: Contém a lógica principal para manipulação de dados do Excel.
- **Vendas.xlsx**: Arquivo de exemplo usado para demonstração.
- **logica.md**: Documentação adicional sobre a lógica implementada.
- **venv/**: Ambiente virtual para gerenciamento de dependências.

## Requisitos
- Python 3.x
- Django
- Pandas
- Bibliotecas adicionais podem ser necessárias dependendo das funcionalidades específicas implementadas.

## Instalação
1. Clone o repositório para o seu ambiente local:
   ```bash
   git clone <URL-do-repositorio>
   ```
2. Navegue até o diretório do projeto:
   ```bash
   cd excelcomdjango
   ```
3. Ative o ambiente virtual:
   ```bash
   source venv/bin/activate  # Para sistemas Unix
   .\venv\Scripts\activate  # Para Windows
   ```
4. Instale as dependências necessárias:
   ```bash
   pip install -r requirements.txt
   ```

## Funcionalidades

- **Importação de Dados**: Importa dados de vendas a partir de um arquivo Excel chamado `Vendas.xlsx`.
- **Cálculo de Faturamento**: Agrupa os dados por `ID Loja` e calcula o faturamento total de cada loja.
- **Cálculo de Quantidade Vendida**: Agrupa os dados por `ID Loja` e calcula a quantidade total de produtos vendidos por loja.
- **Cálculo do Ticket Médio**: Calcula o ticket médio por loja, utilizando a média do valor final e da quantidade de produtos vendidos.
- **Geração de Tabelas HTML**: Cria tabelas em formato HTML para faturamento, quantidade e ticket médio, formatando os valores monetários.
- **Envio de E-mail**: Envia um e-mail com o relatório de vendas por loja, incluindo as tabelas HTML geradas.

## Uso
- Execute o servidor Django:
  ```bash
  python manage.py runserver
  ```
- Acesse a aplicação no navegador através de `http://localhost:8000`.

## Contribuição
Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou enviar pull requests com melhorias ou novas funcionalidades.
