# Projeto de Automação de OnePage

## Habilidades Necessárias:
- "Funções usadas como parâmetro e Expressões Lambda"
- "Análise de Dados com o Pandas + Integração Python e Excel"
- "Integração Python - Arquivos e Pastas do Computador"
- "Integração Python - E-mail"

## Descrição:
Imagine que você trabalha em uma grande rede de lojas de roupas com 25 lojas espalhadas por todo o Brasil. Todos os dias, pela manhã, a equipe de análise de dados calcula os chamados OnePages e envia para o gerente de cada loja o OnePage da sua loja, bem como todas as informações usadas no cálculo dos indicadores.

Um OnePage é um resumo muito simples e direto ao ponto, usado pela equipe de gerência de loja para saber os principais indicadores de cada loja. Ele permite tanto a comparação entre diferentes lojas em uma página quanto a visualização dos indicadores que a loja cumpriu ou não naquele dia.

### Exemplo de OnePage:
```
              Valor Dia   Meta Dia    Cenário Dia
- Indicador1      8           10           X
- Indicador2      7            7           V
- Indicador3      9            5           V
```

Seu papel como Analista de Dados é criar um processo o mais automatizado possível para calcular o OnePage de cada loja e enviar um e-mail para o gerente de cada loja com o OnePage no corpo do e-mail e também o arquivo completo com os dados da sua respectiva loja em anexo.

**Arquivos e Informações Importantes:**
- `Emails.xlsx`: Contém o nome, a loja e o e-mail de cada gerente. Obs: Sugiro substituir a coluna de e-mail de cada gerente por um e-mail seu, para você poder testar o resultado.
- `Vendas.xlsx`: Contém as vendas de todas as lojas. Obs: Cada gerente só deve receber o OnePage e um arquivo em excel em anexo com as vendas da sua loja. As informações de outra loja não devem ser enviadas ao gerente que não é daquela loja.
- `Lojas.csv`: Contém o nome de cada loja.
- Ao final, sua rotina deve enviar ainda um e-mail para a diretoria (informações também estão no arquivo `Emails.xlsx`) com 2 rankings das lojas em anexo, 1 ranking do dia e outro ranking anual. Além disso, no corpo do e-mail, deve ressaltar qual foi a melhor e a pior loja do dia e também a melhor e pior loja do ano. O ranking de uma loja é dado pelo faturamento da loja.
- As planilhas de cada loja devem ser salvas dentro da pasta da loja com a data da planilha, a fim de criar um histórico de Backup.

### Indicadores do OnePage:
- Faturamento -> Meta Ano: 1.650.000 / Meta Dia: 1000
- Diversidade de Produtos (quantos produtos diferentes foram vendidos naquele período) -> Meta Ano: 120 / Meta Dia: 4
- Ticket Médio por Venda -> Meta Ano: 500 / Meta Dia: 500

**Obs:** Cada indicador deve ser calculado no dia e no ano. O indicador do dia deve ser o do último dia disponível na planilha de Vendas (a data mais recente).

Sinta-se à vontade para modificar e expandir este README conforme necessário para a documentação do seu projeto. Boa codificação!