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
import locale
import pprint


# Defina o local desejado (por exemplo, 'pt_BR' para o Brasil)
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

#FUNÇÃO -> Escrever o email:
def enviar_email(loja, dia_indicador, arquivo_anexo, faturamento_dia, faturamento_ano, qtde_dia, qtde_ano, ticket_medio_dia, ticket_medio_ano):
    # Definição de metas
    meta_faturamento_dia = 1000
    meta_faturamento_ano = 1650000
    meta_qtdeprodutos_dia = 4
    meta_qtdeprodutos_ano = 120
    meta_ticketmedio_dia = 500
    meta_ticketmedio_ano = 500

    # Obter o nome e o email do gerente da loja
    nome = ''
    email_gerente = ''
    gerente_loja = emails_df[emails_df['Loja'] == loja]
    if not gerente_loja.empty:
        nome = gerente_loja['Gerente'].values[0]
        email_gerente = gerente_loja['E-mail'].values[0]
    else:
        print(f"Não há informações de gerente para a loja {loja}.")

    if faturamento_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'

    if faturamento_ano >= meta_faturamento_ano:
        cor_fat_ano = 'green'
    else:
        cor_fat_ano = 'red'

    if qtde_dia >= meta_qtdeprodutos_dia:
        cor_qtde_dia = 'green'
    else:
        cor_qtde_dia = 'red'

    if qtde_ano >= meta_qtdeprodutos_ano:
        cor_qtde_ano = 'green'
    else:
        cor_qtde_ano = 'red'

    if ticket_medio_dia >= meta_ticketmedio_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'

    if ticket_medio_ano >= meta_ticketmedio_ano:
        cor_ticket_ano = 'green'
    else:
        cor_ticket_ano = 'red'

    # Criar o corpo do email
    corpo_email = f"""
    <p>Bom dia, {nome}</p>
    <p>O resultado de ontem <strong>({dia_indicador.day}/{dia_indicador.month})</strong> da <strong>Loja {loja}</strong>, foi: </p>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Dia</th>
        <th>Meta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style='text-align:center'>R${faturamento_dia:.2f}</td>
        <td style='text-align:center'>R${meta_faturamento_dia:.2f}</td>
        <td style='text-align:center'><font color='{cor_fat_dia}'>■</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style='text-align:center'>{qtde_dia}</td>
        <td style='text-align:center'>{meta_qtdeprodutos_dia}</td>
        <td style='text-align:center'><font color='{cor_qtde_dia}'>■</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style='text-align:center'>R${ticket_medio_dia:.2f}</td>
        <td style='text-align:center'>R${meta_ticketmedio_dia:.2f}</td>
        <td style='text-align:center'><font color='{cor_ticket_dia}'>■</font></td>
      </tr>
    </table>
    <br>
    <table>
      <tr>
        <th>Indicador</th>
        <th>Valor Ano</th>
        <th>Meta Ano</th>
        <th>Cenário Ano</th>
      </tr>
      <tr>
        <td>Faturamento</td>
        <td style='text-align:center'>R${faturamento_ano:.2f}</td>
        <td style='text-align:center'>R${meta_faturamento_ano:.2f}</td>
        <td style='text-align:center'><font color='{cor_fat_ano}'>■</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style='text-align:center'>{qtde_ano}</td>
        <td style='text-align:center'>{meta_qtdeprodutos_ano}</td>
        <td style='text-align:center'><font color='{cor_qtde_ano}'>■</font></td>
      </tr>
      <tr>
        <td>Ticket Médio</td>
        <td style='text-align:center'>R${ticket_medio_ano:.2f}</td>
        <td style='text-align:center'>R${meta_ticketmedio_ano:.2f}</td>
        <td style='text-align:center'><font color='{cor_ticket_ano}'>■</font></td>
      </tr>
    </table>
    <p>Segue em anexo a planilha com todos os dados para mais detalhes. </p>
    <p>Qualquer dúvida estou à disposição. </p>
    <p>Att., Bernardo </p>
    """

    # Configurar a mensagem do email
    msg = MIMEMultipart()
    msg['Subject'] = f"OnePage dia {dia_indicador.day}/{dia_indicador.month} - Loja {loja}"
    msg['From'] = 'be7ferrao@gmail.com'  # Coloque seu email aqui
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
    password = 'unrn okku guop cqsa'  # Coloque sua senha aqui
    s.login(msg['From'], password)

    # Enviar o email
    s.sendmail(msg['From'], msg['To'], msg.as_string())
    print('Email enviado')
def formatar_valor(valor):
    # Remover "R$" e espaços em branco
    valor_limpo = valor.replace('R$', '').replace(' ', '')

    # Substituir vírgulas por pontos e remover ponto de milhar
    valor_limpo = valor_limpo.replace(',', '.').replace('.', '', 1)

    # Retornar apenas a parte numérica
    return float(valor_limpo)


#1 - Mandar a Onepage
#(Carregar os dados):
# Obter o caminho do diretório atual usando pathlib
caminho = Path.cwd() / 'Arquivos'
# Carregar os DataFrames
vendas_df = pd.read_excel(caminho / 'vendas.xlsx')
vendas_df['Data'] = pd.to_datetime(vendas_df['Data'])
vendas_df['Valor Unitário'] = vendas_df['Valor Unitário'].apply(formatar_valor)
vendas_df['Valor Final'] = vendas_df['Valor Final'].apply(formatar_valor)
emails_df = pd.read_excel(caminho / 'Emails.xlsx')
lojas_df = pd.read_excel(caminho / 'Lojas.xlsx')

#(Tratamento dos dados) - Unir os dataframes Vendas e Lojas através de um merge(Coluna ID Loja):
vendas_df = pd.merge(vendas_df, lojas_df, on='ID Loja')
# Exibir a tabela formatada usando tabulate
#print(tabulate(merged_df.head(5), headers='keys', tablefmt='fancy_grid'))

##(Tratamento dos dados) - Criar um df para cada loja:
dicionario_lojas = {} #Dicionario de tabelas/dfs de lojas
for loja in lojas_df['Loja']: #Passa por todos os nomes da loja
    dicionario_lojas[loja] = vendas_df.loc[vendas_df['Loja'] == loja, :] #Pega todas as linhas de uma loja específica
pprint.pprint(dicionario_lojas)

###(Tratamento dos dados) - Dia do indicador(Calcular o último dia presente) e último ano:
dia_indicador = vendas_df['Data'].max()
# print(f'{dia_indicador.day}/{dia_indicador.month}')

#2 - Fazer Backup(Criar pastas com o nome das lojas, salvar o arquivo de cada loja)
#Identificar se as pastas estão criadas
caminho_backup = pathlib.Path(r'Arquivos/Backup')
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
# Enviar emails para cada loja
for loja in dicionario_lojas:
    vendas_loja = dicionario_lojas[loja]
    vendas_loja_dia = vendas_loja.loc[vendas_loja['Data']==dia_indicador, :]

    #faturamento
    faturamento_ano = vendas_loja['Valor Final'].sum()
    #print(faturamento_ano)
    faturamento_dia = vendas_loja_dia['Valor Final'].sum()
    #print(faturamento_dia)

    #diversidade de produtos
    qtde_ano = len(vendas_loja['Produto'].unique())
    #print(qtde_produtos_ano)
    qtde_dia = len(vendas_loja_dia['Produto'].unique())
    #print(qtde_produtos_dia)

    #ticket medio
    valor_venda = vendas_loja.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_ano = valor_venda['Valor Final'].mean()

    valor_venda_dia = vendas_loja_dia.groupby('Código Venda').sum(numeric_only=True)
    ticket_medio_dia = valor_venda_dia['Valor Final'].mean()

    attachment = caminho / 'Backup' / loja / f'{dia_indicador.month}_{dia_indicador.day}_{loja}.xlsx'
    enviar_email(loja, dia_indicador, attachment, faturamento_dia, faturamento_ano, qtde_dia, qtde_ano, ticket_medio_dia, ticket_medio_ano)

#Criar o ranking das lojas pelo faturamento
faturamento_lojas = vendas_df.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_ano = faturamento_lojas.sort_values(by='Valor Final', ascending=False)

vendas_dia = vendas_df.loc[vendas_df['Data'] == dia_indicador, :]
faturamento_lojas_dia = vendas_dia.groupby('Loja')[['Loja', 'Valor Final']].sum()
faturamento_lojas_dia = faturamento_lojas_dia.sort_values(by='Valor Final', ascending=False)

# Caso esteja criada, salvar o arquivo
nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
faturamento_lojas_ano.to_excel(rf'C:\Users\be-8f\PycharmProjects\Process-Automation-Project\Arquivos\Backup\{nome_arquivo}', index=False)

nome_arquivo = f'{dia_indicador.month}_{dia_indicador.day}_Ranking Diario.xlsx'
faturamento_lojas_dia.to_excel(rf'C:\Users\be-8f\PycharmProjects\Process-Automation-Project\Arquivos\Backup\{nome_arquivo}', index=False)

def enviar_ranking(arquivo1, arquivo2, faturamento_lojas_dia, faturamento_lojas_ano):
    email_gerente = emails_df.loc[emails_df['Loja'] == 'Diretoria', 'E-mail'].values[0]

    # Criar o corpo do email
    corpo_email = f"""
    Prezados, bom dia

    Melhor loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[0]} com Faturamento R${faturamento_lojas_dia.iloc[0, 0]:.2f}
    Pior loja do Dia em Faturamento: Loja {faturamento_lojas_dia.index[-1]} com Faturamento R${faturamento_lojas_dia.iloc[-1, 0]:.2f}

    Melhor loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[0]} com Faturamento R${faturamento_lojas_ano.iloc[0, 0]:.2f}
    Pior loja do Ano em Faturamento: Loja {faturamento_lojas_ano.index[-1]} com Faturamento R${faturamento_lojas_ano.iloc[-1, 0]:.2f}

    Segue em anexo os rankings do ano e do dia de todas as lojas.

    Qualquer dúvida estou à disposição.

    Att.,
    Bernardo      
    """

    # Configurar a mensagem do email
    msg = MIMEMultipart()
    msg['Subject'] = f'Ranking Dia {dia_indicador.day}/{dia_indicador.month}'
    msg['From'] = 'be7ferrao@gmail.com'  # Coloque seu email aqui
    msg['To'] = email_gerente

    # Adicionar o corpo do email
    msg.attach(MIMEText(corpo_email, 'html'))

    # Adicionar o primeiro anexo ao email
    with open(arquivo1, 'rb') as attachment1:
        part1 = MIMEBase('application', 'octet-stream')
        part1.set_payload(attachment1.read())
        encoders.encode_base64(part1)
        part1.add_header('Content-Disposition', f'attachment; filename= {arquivo1.name}')
        msg.attach(part1)

    # Adicionar o segundo anexo ao email
    with open(arquivo2, 'rb') as attachment2:
        part2 = MIMEBase('application', 'octet-stream')
        part2.set_payload(attachment2.read())
        encoders.encode_base64(part2)
        part2.add_header('Content-Disposition', f'attachment; filename= {arquivo2.name}')
        msg.attach(part2)

    # Configurações de segurança
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    password = 'unrn okku guop cqsa'  # Coloque sua senha aqui
    s.login(msg['From'], password)

    # Enviar o email
    s.sendmail(msg['From'], msg['To'], msg.as_string())
    print('Email enviado')

# Exemplo de uso
arquivo1 = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Anual.xlsx'
arquivo2 = pathlib.Path.cwd() / caminho_backup / f'{dia_indicador.month}_{dia_indicador.day}_Ranking Diario.xlsx'
enviar_ranking(arquivo1, arquivo2, faturamento_lojas_dia, faturamento_lojas_ano)
