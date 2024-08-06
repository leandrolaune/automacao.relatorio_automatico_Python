# Descrição
# Projeto feito por Leandro Sanches Launé
# project made by Leandro Sanches Launé

# Description
**Portuguese Version:**  
Esse projeto aborda a construção de uma automação em python do processo de extração de uma base de dados do excel, análise de faturamento e manipulação de dados utilizando Pandas e envio de email com tabela contendo relatório do faturamento total, obtida automaticamente a partir da base de dados juntamnte com o uso do código. 
**English Version:**  
This project addresses the construction of an automation in python of the process of extracting an excel database, billing analysis and data manipulation using Pandas and sending an email with a table containing a report of the total revenue, automatically obtained from the database along with the use of the code. 
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
### Resultado
### Result
![envio email1](https://github.com/user-attachments/assets/719c3f35-97de-4fd9-a38d-4acf80018e38)
![envio email2](https://github.com/user-attachments/assets/4ddd5a8b-f459-41ea-b9e8-508b8521fc66)

