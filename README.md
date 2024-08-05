# Descrição
# Description
**Portuguese Version:**  
Esse projeto aborda a construção de uma automação em python do processo de importação de uma base de dados do excel e envio de email com tabela contendo relatório do faturamento total, obtido automaticamente a partir da base de dados com o auxílio do código da automação
Tal projeto foi desenvolvido utilizando a biblioteca pandas para manipulação adequada dos dados, para dessa forma, viabilizar a construção automática da tabela contendo o relatório.  
**English Version:**  
Esse projeto aborda a construção de uma automação em python do processo de importação de uma base de dados do excel e envio de email com tabela contendo relatório do faturamento total, obtido automaticamente a partir da base de dados com o auxílio do código da automação
Tal projeto foi desenvolvido utilizando a biblioteca pandas para manipulação adequada dos dados, para dessa forma, viabilizar a construção automática da tabela contendo o relatório.  
### Importando Bibliotecas:
### Importing Libraries:

```
import pandas as pd
import win32com.client as win32
```


### importar a base de dados
### import the database

```
tabela_vendas = pd.read_excel(r"C:\Users\leand\Downloads\Vendas.xlsx")
```

### visualizar a base de dados
### view the database

```
pd.set_option('display.max_columns', None)
print(tabela_vendas)
```

### faturamento por loja
### revenue per store

```
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby(
    'ID Loja').sum()
print(faturamento)
```

### quantidade de produtos vendidos por loja
### number of products sold per store

```
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
```

### ticket médio por produto em cada loja
### average ticket per product in each store

```
ticket_medio = (faturamento['Valor Final'] /
                quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)
```

### enviar um email com o relatório
### send an email with the report

```
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'leandroabc1401@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estamos à disposição.</p>

<p>Att.,</p>
<p>SanData</p>
'''

mail.Send()

print('Email Enviado')
```
