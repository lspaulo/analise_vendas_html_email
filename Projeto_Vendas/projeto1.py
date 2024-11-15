# Importação das bibliotecas
import pandas as pd
import plotly.express as px
import smtplib
import email.message as em
from babel.numbers import format_currency

# Passo 1 - Importar a base de dados (Pasta vendas)

tabela_vendas = pd.read_excel("/content/drive/MyDrive/Phyton para Ciência de Dados/Projeto 1/Vendas/Vendas+-+Base+de+Dados.xlsx")
display(tabela_vendas)

# Passo 2 - Calular o produto mais vendido(em quantidade)
# Agrupando por produto e somando as linhas referente ao produto
tb_quantidade_produto = tabela_vendas.groupby('Produto').sum()
# Filtrando colunas específicas, neste caso foi a coluna quantidade
tb_quantidade_produto = tb_quantidade_produto[['Quantidade']]
# tb_quantidade_produto = tb_quantidade_produto.sort_values(by='Quantidade', ascending=False)
display(tb_quantidade_produto)



# Passo 3 - Calcular o Produto mais vendido(em faturamento).
tabela_vendas['Faturamento'] = tabela_vendas['Quantidade'] * tabela_vendas['Valor Unitário']
tb_faturamento_produto = tabela_vendas.groupby('Produto').sum()
tb_faturamento_produto = tb_faturamento_produto[['Faturamento']].sort_values(by='Faturamento', ascending=False)
display(tb_faturamento_produto)

# Passo 4 - Calcular loja/estado que mais vendeu(em faturamento)
tb_faturamento_loja = tabela_vendas.groupby('Loja').sum()
tb_faturamento_loja = tb_faturamento_loja[['Faturamento']].sort_values(by='Faturamento', ascending=False)
display(tb_faturamento_loja)

# Passo 5 - Calular o ticket médio por loja/estado
tabela_vendas['Ticket Médio'] = tabela_vendas['Valor Unitário']
tb_ticket_medio = tabela_vendas.groupby('Loja').mean(numeric_only=True)
tb_ticket_medio = tb_ticket_medio[['Ticket Médio']]
tb_ticket_medio = tb_ticket_medio.sort_values(by='Ticket Médio', ascending=False)
display(tb_ticket_medio)

# Passo 6 - Criar uma gráfico/dashboard da loja/estado qua mais vendeu(em faturamento)

grafico = px.bar(tb_faturamento_loja, y='Faturamento', x=tb_faturamento_loja.index)
grafico.show()

# Formatação monetária
tb_faturamento_loja_formatado = pd.DataFrame(tb_faturamento_loja['Faturamento'].apply(lambda x: format_currency(x,'BRL', locale='pt_BR')))
tb_faturamento_loja_formatado = tb_faturamento_loja_formatado.reset_index()

# Formatação HTML
tb_faturamento_loja_formatado_html = tb_faturamento_loja_formatado.to_html(index=False, justify='center', border=0).replace('<tbody>', '<tbody style="text-align: center; color: #484848; background: #f7f7f7">')

display(tb_faturamento_loja_formatado)
display(tb_faturamento_loja_formatado_html)

# Formatação númerica

tb_quantidade_produto_formatado = pd.DataFrame(tb_quantidade_produto['Quantidade'].apply(lambda x: "{:,}".format(x).replace(',', '.')))
tb_quantidade_produto_formatado = tb_quantidade_produto_formatado.reset_index()

# Formatação HTML
tb_quantidade_produto_formatado_html = tb_quantidade_produto_formatado.to_html(index=False, justify='center', border=0).replace('<tbody>', '<tbody style="text-align: center; color: #484848; background: #f7f7f7">')


display(tb_quantidade_produto_formatado)
display(tb_quantidade_produto_formatado_html)

# Formatação monetária
tb_faturamento_produto_formatado = pd.DataFrame(tb_faturamento_produto['Faturamento'].apply(lambda x: format_currency(x,'BRL', locale='pt_BR')))
tb_faturamento_produto_formatado = tb_faturamento_produto_formatado.reset_index()

# Formatação HTML
tb_faturamento_produto_formatado_html = tb_faturamento_produto_formatado.to_html(index=False, justify='center', border=0).replace('<tbody>', '<tbody style="text-align: center; color: #484848; background: #f7f7f7">')

display(tb_faturamento_produto_formatado)
display(tb_faturamento_produto_formatado_html)

# Formatação monetária
tb_ticket_medio_formatado = pd.DataFrame(tb_ticket_medio['Ticket Médio'].apply(lambda x: format_currency(x,'BRL', locale='pt_BR')))
tb_ticket_medio_formatado = tb_ticket_medio_formatado.reset_index()

# Formatação HTML

tb_ticket_medio_formatado_html = tb_ticket_medio_formatado.to_html(index=False, justify='center', border=0).replace('<tbody>', '<tbody style="text-align: center; color: #484848; background: #f7f7f7">')

display(tb_ticket_medio_formatado)
display(tb_ticket_medio_formatado_html)

"""**Enviar relatórios de indicadores de desempenho**"""

# Passo 7 - Enviar um e-mail para o setor responsável.

corpo_email = f"""
<p>Prezados, tudo bem?</p>

<p>Segue o relatórios de vendas do mês.</p>

<p><b>Faturamento por loja:</b></p>
<p>{tb_faturamento_loja_formatado_html}</p>

<p><b>Quantidade vendida por produto:</b></p>
<p>{tb_quantidade_produto_formatado_html}</p>

<p><b>Faturamento por produto:</b></p>
<p>{tb_faturamento_produto_formatado_html}</p>

<p><b>Ticket medio por loja:</b></p>
<p>{tb_ticket_medio_formatado_html}</p>

<p>Qualquer dúvida estou a disposição.</p>

<p>Att.</p>
<p>Paulo</>
"""

#settings to send
msg = em.Message()
msg['Subject'] = "Relatório de Vendas!" #ASSUNTO DO E-MAIL#
msg['From'] = 'contatolspl@gmail.com' #E-MAIL QUE VAI ENVIAR O E-MAIL#
msg['To'] = 'luispaulo9919@gmail.com'#E-MAIL QUE VAI RECEBER
password = '********' #SENHA DO E-MAIL QUE VAI ENVIAR(tem q gerar no gmail)
msg.add_header('Content-Type', 'text/html') # Tipo/fomato do conteúdo
msg.set_payload(corpo_email )# Mensagem do corpo do e-mail que está na variávelcorpo_email

#settings gmail server
s = smtplib.SMTP('smtp.gmail.com: 587')
s.starttls()

# Login Credentials for sending the mail
s.login(msg['From'], password)
s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
print('Email enviado')