# 📊 Automação de Relatório de Vendas com Envio por E-mail

Este projeto automatiza a consolidação de arquivos CSV contendo dados de vendas, trata as datas no formato Excel, gera um relatório em Excel e envia por e-mail via Outlook.

## 🚀 Funcionalidades

- Leitura automática de múltiplos arquivos CSV da pasta `./bases`
- Conversão da coluna `Data de Venda` do formato Excel para data legível
- Consolidação e ordenação dos dados
- Geração de um arquivo Excel (`vendas.xlsx`)
- Envio automático de e-mail com o relatório em anexo

## 🛠️ Requisitos

- Python 3.x
- Pacotes:
  - `pandas`
  - `pywin32` (para integração com Outlook)
- Microsoft Outlook instalado e configurado

## 📂 Estrutura de Pastas
```plaintext
📁 projeto/
├── 📁 bases/
│   ├── vendas_janeiro.csv
│   ├── vendas_fevereiro.csv
│   └── ...
├── vendas.xlsx
├── app.py
└── README.md
```

## 📧 Configuração do E-mail
No script, edite a linha abaixo com os destinatários desejados:
```python
email.To = 'exemplo@gmail.com; exemplo2@outlook.com.br'
```
## 🕒 Agendamento (opcional)
```
Para executar automaticamente todos os dias:

1. Abra o Agendador de Tarefas do Windows
2. Crie uma nova tarefa
3. Configure o gatilho (ex: diariamente às 8h)
4. Na ação, selecione:
  Programa/script: python
  Adicionar argumentos: caminho\para\script.py
```
## ✅ Como Executar
1. Coloque os arquivos CSV na pasta ./bases

2. Execute o script:
```bash
  python app.py
```
3. O arquivo vendas.xlsx será gerado e enviado por e-mail automaticamente

## ⚠️ Observações
```
Certifique-se de que o Outlook esteja aberto ou configurado corretamente para envio

A data no Excel pode ter um deslocamento de 1 dia por conta de um bug histórico (ano 1900 como bissexto). Se necessário, ajuste a base para '1899-12-30' no cálculo de datas.
```
## ✍️ Autor
```
Adones Melo 
💼 Automação de processos com Python
📧 adones.n.m@outlook.com
```




