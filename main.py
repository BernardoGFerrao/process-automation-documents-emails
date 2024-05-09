#Bibliotecas:
import pandas as pd
import pathlib #Funciona não somento no windows
from pathlib import Path
from tabulate import tabulate
import smtplib
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#1 - Mandar a Onepage
#(Carregar os dados):
# Obter o caminho do diretório atual usando pathlib
caminho = Path.cwd() / 'Arquivos'
# Carregar os DataFrames
vendas_df = pd.read_excel(caminho / 'vendas.xlsx')
vendas_df['Data'] = pd.to_datetime(vendas_df['Data'])
emails_df = pd.read_excel(caminho / 'Emails.xlsx')
lojas_df = pd.read_excel(caminho / 'Lojas.xlsx')

#(Tratamento dos dados) - Unir os dataframes Vendas e Lojas através de um merge(Coluna ID Loja):
merged_df = pd.merge(vendas_df, lojas_df, on='ID Loja')
# Exibir a tabela formatada usando tabulate
#print(tabulate(merged_df.head(5), headers='keys', tablefmt='fancy_grid'))

##(Tratamento dos dados) - Criar um df para cada loja:
dicionario_lojas = {} #Dicionario de tabelas/dfs de lojas
for loja in lojas_df['Loja']: #Passa por todos os nomes da loja
    dicionario_lojas[loja] = merged_df.loc[merged_df['Loja'] == loja, :] #Pega todas as linhas de uma loja específica
print(dicionario_lojas)

###(Tratamento dos dados) - Dia do indicador(Calcular o último dia presente) e último ano:
dia_indicador = merged_df['Data'].max()
# print(f'{dia_indicador.day}/{dia_indicador.month}')

#2 - Fazer Backup(Criar pastas com o nome das lojas, salvar o arquivo de cada loja)
#Identificar se as pastas estão criadas
caminho_backup = pathlib.Path(r'Backup')
arquivos_pasta_backup = caminho_backup.iterdir()#Lista de todos arquivos dentro da pasta backup
lista_arquivos_pasta_backup = [arquivo.name for arquivo in arquivos_pasta_backup]#Lista criada para armazenar o NOME dos arquivos dentro da pasta backup
#Salvar dentro da pasta
for loja in dicionario_lojas:
    if loja not in lista_arquivos_pasta_backup:
    #Criar a pasta com o nome da loja:
        nova_pasta = caminho_backup / loja
        nova_pasta.mkdir(parents=True, exist_ok=True)
    #Caso esteja criada, salvar o arquivo
    nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    local_arquivo = caminho_backup / loja / nome_arquivo #O nome do arquivo será o nome que ficará ao ser criado
    dicionario_lojas[loja].to_excel(local_arquivo, index=False)

#3 - Enviar resumo para diretória
#Criação dos indicadores
for loja in dicionario_lojas:
    ##Faturamento:
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data'] == dia_indicador, :]
    faturamento_ano = vendas_loja['Valor Final'].sum()
    faturamento_dia = vendas_loja['Valor Final'].sum()
    # print(faturamento_ano)
    # print(faturamento_dia)

    ##Diversidade de produtos:
    qtde_ano = vendas_loja['Produto'].nunique()
    qtde_dia = vendas_loja_dia['Produto'].nunique()
    # print(qtde_ano)
    # print(qtde_dia)

    ## ticket medio
    vendas_loja.drop('Data', axis=1, inplace=True)
    valor_venda = vendas_loja.groupby('Código Venda').sum()
    ticket_medio_ano = valor_venda['Valor Final'].mean()
    # print(ticket_medio_ano)
    # ticket_medio_dia
    vendas_loja_dia.drop('Data', axis=1, inplace=True)
    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum()
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()
    # print(ticket_medio_dia)

#Definição de metas:
meta_dia = 1000
meta_ano = 1650000
meta_qtdeprodutos_dia = 4
meta_qtdeprodutos_ano = 120
meta_ticketmedio_dia = 500
meta_ticketmedio_ano = 500

#Escrever o email:
def enviar_email(loja, dia_indicador, arquivo_anexo):
    # Criar o corpo do email
    corpo_email = """
    <p>Parágrafo1</p>
    <p>Parágrafo2</p>
    """

    # Obter o nome e o email do gerente da loja
    nome = emails_df.loc[emails_df['Loja'] == loja, 'Gerente'].values[0]
    email_gerente = emails_df.loc[emails_df['Loja'] == loja, 'E-mail'].values[0]

    # Configurar a mensagem do email
    msg = MIMEMultipart()
    msg['Subject'] = f"OnePage dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}"
    msg['From'] = 'meuemail'
    msg['To'] = email_gerente

    # Adicionar o corpo do email
    msg.attach(MIMEText(corpo_email, 'html'))

    # Adicionar o anexo ao email
    with open(arquivo_anexo, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())

    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {arquivo_anexo}')
    msg.attach(part)

    # Configurações de segurança
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    password = 'minhasenha'  # Coloque sua senha aqui
    s.login(msg['From'], password)

    # Enviar o email
    s.sendmail(msg['From'], msg['To'], msg.as_string())
    print('Email enviado')


# Exemplo de uso
# for loja in dicionario_lojas:
#     enviar_email(loja, dia_indicador, 'caminho/para/o/arquivo.pdf')
