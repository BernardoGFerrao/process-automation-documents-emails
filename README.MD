**Portuguese Section:**
# Projeto de Automação de OnePage

Este projeto de automação foi desenvolvido para calcular e enviar o OnePage diário de cada loja em uma grande rede de lojas de roupas no Brasil. Utiliza habilidades como funções como parâmetro, expressões lambda, análise de dados com Pandas, integração Python e Excel, integração Python com arquivos e pastas do computador, e integração Python com e-mail.

## Funcionalidades

- **Cálculo do OnePage:** Calcula os indicadores do OnePage (Faturamento, Diversidade de Produtos, Ticket Médio por Venda) para cada loja.
- **Envio de E-mails:** Envia o OnePage para o gerente de cada loja, além de um e-mail para a diretoria com rankings das lojas.
- **Backup de Planilhas:** Salva as planilhas de cada loja com a data, criando um histórico de backup.

## Como Funciona

Este projeto automatiza o cálculo do OnePage de cada loja e o envio por e-mail para o gerente correspondente, seguindo as seguintes etapas:
1. **Cálculo dos Indicadores:** Utiliza dados de vendas para calcular os indicadores do OnePage (Faturamento, Diversidade de Produtos, Ticket Médio por Venda) para cada loja.
2. **Envio de E-mails:** Envia o OnePage para cada gerente, juntamente com um arquivo Excel das vendas da respectiva loja em anexo.
3. **Rankings para a Diretoria:** Envia um e-mail para a diretoria com rankings diário e anual das lojas, destacando a melhor e pior loja do dia e do ano.

## Arquivos Necessários

- **Emails.xlsx:** Contém informações dos gerentes, incluindo nome, loja e e-mail (sugere-se substituir por um e-mail de teste).
- **Vendas.xlsx:** Contém dados de vendas de todas as lojas.
- **Lojas.csv:** Lista o nome de cada loja.

## Indicadores do OnePage

- **Faturamento:** Meta Ano: 1.650.000 / Meta Dia: 1000
- **Diversidade de Produtos:** Meta Ano: 120 / Meta Dia: 4
- **Ticket Médio por Venda:** Meta Ano: 500 / Meta Dia: 500

---

**English Section:**

# OnePage Automation Project

This automation project was developed to calculate and send the daily OnePage of each store in a large clothing store network in Brazil. It utilizes skills such as functions as parameters, lambda expressions, data analysis with Pandas, Python and Excel integration, Python integration with computer files and folders, and Python integration with email.

## Features

- **OnePage Calculation:** Calculates the OnePage indicators (Revenue, Product Diversity, Average Ticket per Sale) for each store.
- **Email Sending:** Sends the OnePage to each store manager, along with an email to the board with store rankings.
- **Backup of Spreadsheets:** Saves each store's spreadsheets with the date, creating a backup history.

## How It Works

This project automates the calculation of each store's OnePage and sends it by email to the corresponding manager, following these steps:
1. **Calculation of Indicators:** Uses sales data to calculate the OnePage indicators (Revenue, Product Diversity, Average Ticket per Sale) for each store.
2. **Email Sending:** Sends the OnePage to each manager, along with an Excel file of the store's sales attached.
3. **Rankings for the Board:** Sends an email to the board with daily and annual rankings of the stores, highlighting the best and worst store of the day and year.

## Required Files

- **Emails.xlsx:** Contains manager information, including name, store, and email (it is suggested to replace with a test email).
- **Sales.xlsx:** Contains sales data for all stores.
- **Stores.csv:** Lists the name of each store.

## OnePage Indicators

- **Revenue:** Year Goal: 1,650,000 / Daily Goal: 1000
- **Product Diversity:** Year Goal: 120 / Daily Goal: 4
- **Average Ticket per Sale:** Year Goal: 500 / Daily Goal: 500