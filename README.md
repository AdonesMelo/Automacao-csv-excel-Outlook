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
├── script.py
└── README.md

