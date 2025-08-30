import os
from datetime import datetime
import pandas as pd
import win32com.client as win32

caminho = './bases'
arquivos = os.listdir(caminho)

tabela_consolidada = pd.DataFrame()

# Trata a coluna 'Data de Venda', porque as data estão no formato do excel de dias
for nome_arquivo in arquivos:
    tabela_vendas = pd.read_csv(os.path.join(caminho, nome_arquivo))
    tabela_vendas['Data de Venda'] = pd.to_datetime('01/01/1900') + pd.to_timedelta(tabela_vendas['Data de Venda'], unit='d')
    tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])

# Ordenar a colana 'Data de Venda'
tabela_consolidada = tabela_consolidada.sort_values(by='Data de Venda')
tabela_consolidada = tabela_consolidada.reset_index(drop=True) # Resetar o index

tabela_consolidada['Data de Venda'] = tabela_consolidada['Data de Venda'].dt.date # remover a hora
tabela_consolidada.to_excel('vendas.xlsx', index=False) # cria o arquivo excel
print('Arquivo excel criado com sucesso!')

# conexão com outlook
outlook = win32.Dispatch('outlook.application')
email = outlook.CreateItem(0)
email.To = 'exemplo@gmail.com; exemplo2@outlook.com.br' # passa os contatos dentro da string separado por ;
data_hj = datetime.today().strftime('%d/%m/%Y')
email.Subject = f'Relatórios das vendas {data_hj}' 
email.Body = f'''
Presados,

Segue em anexo o relatório de vendas do dia {data_hj} atualizado.
Qualquer coisa estou á disposição

abs,
Adones Melo
'''

caminho = os.getcwd()
anexo = os.path.join(caminho, 'vendas.xlsx')
email.Attachments.Add(anexo)

email.Send()
print('E-mails enviado com sucesso!')