import csv
import os.path
import ofxparse
import support
import datetime

# Exemplos
# https://python.hotexamples.com/pt/examples/ofxparse.ofxparse/OfxParser/-/python-ofxparser-class-examples.html
# https://programtalk.com/vs2/python/6616/ofxparse/tests/test_parse.py/
ofx = ofxparse.OfxParser.parse(support.open_file('file.ofx'))

# AccountType
# (Unknown, Bank, CreditCard, Investment)

# Account

account = ofx.account
print(account.account_id)  # The account number
print(account.number)  # The account number (deprecated -- returns account_id)
print(account.routing_number)  # The bank routing number
account.branch_id  # Transit ID / branch number
print(account.type)  # An AccountType object
account.statement  # A Statement object
account.institution  # An Institution object

# InvestmentAccount(Account)

# account.brokerid          # Investment broker ID
account.statement  # An InvestmentStatement object

# Institution

institution = account.institution
print(institution.organization)
print(institution.fid)

# Statement

statement = account.statement
statement.start_date  # The start date of the transactions
statement.end_date  # The end date of the transactions
statement.balance  # The money in the account as of the statement date
# statement.available_balance   # The money available from the account as of the statement date
statement.transactions  # A list of Transaction objects

# InvestmentStatement

statement = account.statement
# statement.positions           # A list of Position objects
statement.transactions  # A list of InvestmentTransaction objects

# Transaction
# O cabecalho da planilha eh o seguinte
"""Nome | Saldo  atual | Conta | Transferencia | Descricao | Beneficiario | Categoria | Data | Tempo | Valor | Cambio | Numero do cheque | Saldo"""

rows = []
for transaction in statement.transactions:
    # print(date_time)
    rows.append(('', '',  institution.organization + " " + account.number, '',
                 transaction.memo, '', '', transaction.date,
                 transaction.amount, '', 'BRL', ''))

    #print(transaction.payee)
    #print(transaction.type)
    #print(transaction.date)
    # transaction.user_date
    #print(transaction.amount)
    #print(transaction.id)
    #print(transaction.memo)
    #print(transaction.sic)
    #print(transaction.mcc)
    #print(transaction.checknum)

for i in rows:
    print(i)

if os.path.exists('dados.csv'):
    with open('dados.csv', 'a', newline='', encoding='utf-8') as arq:
        dados = csv.writer(arq, lineterminator='\n', delimiter=';')
        dados.writerows(rows)
else:
    with open('dados.csv', 'w') as arq:
        dados = csv.writer(arq, lineterminator='\n', delimiter=';')
        dados.writerow(
            ['Nome', 'Saldo atual', 'Conta', 'Transferências', 'Descrição', 'Beneficiário', 'Categoria',
             'Data', 'Tempo', 'Valor', 'Câmbio', 'Número do cheque', 'Saldo'])
        dados.writerows(rows)
