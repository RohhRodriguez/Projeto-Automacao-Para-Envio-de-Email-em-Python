#Biblioteca que permite integração do PY com excel (instação com o comando "pip install pandas")
#Biblioteca que permite integração do PY com Outlook (instalação com o comando "pip install pywin32"
import pandas as pd
import win32com.client as win32

#importar a base de dados
tabela_vendas = pd.read_excel('vendas.xlsx')

#visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

#faturamento por loja
#tabela_vendas[['ID Loja','Valor Final']] => para filtrar resultado apenas das colunas que deseja que apareça na tabela
#.groupby('ID Loja') => para agrupar todas as lojas e mostrar uma linha para cada
#.sum() => para somar valores da outra coluna filtrada, no caso, 'Valor Final'. O resultado será uma exibição de faturamento total por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

#quantidade de produtos vendidos
qtdevendida = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtdevendida)

print('-' * 50)


#ticket médio por loja
#ticketmedio = faturamento/qtdevendida
#.to_frame => transformando uma série de dados em uma tabela
ticketMedio = (faturamento['Valor Final'] / qtdevendida['Quantidade']).to_frame()
ticketMedio = ticketMedio.rename(columns={0: 'Ticket Médio'})
print(ticketMedio)


#enviar um e-mail com o relatório:

#código original para envio de email com outlook(modelo internet):
#outlook = win32.Dispatch('outlook.application')
#mail = outlook.CreateItem(0)
#mail.To = 'To address'
#mail.Subject = 'Message subject'
#mail.HTMLBody = '<h2>HTML Message body</h2>'
#mail.Send()
#.to_html() => transforma o retorno em uma tabela dentro do HTML
#R$1.000.000,00 => para formatar como moeda, passar "(formatters={'Valor Final': 'R${:,.2f}'.format})")

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'rohh3578@gmail.com'
mail.Subject = 'Relatório diário de vendas'
mail.HTMLBody = f'''
<p>Prezados,</p>
<p>Segue relatório diário de vendas por loja.</p>
<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{qtdevendida.to_html()}

<p>Ticket médio dos produtos vendidos em cada loja:</p>
{ticketMedio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida, fico à disposição!</p>

<p>Att,</p>
<p>Rônica CS Rodrigues</p>
'''
mail.Send()

print('Email Enviado')

