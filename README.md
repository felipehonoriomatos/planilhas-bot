# Automação de Planilhas Excel e Google Sheets com Python

Automação completa para leitura, processamento e geração de relatórios
em Excel (.xlsx) e Google Sheets — sem abrir nenhum programa manualmente.

---

## Funcionalidades

-  Lê planilhas Excel e gera relatório formatado automaticamente
-  Calcula totais, médias, maiores e menores valores por categoria
-  Destaca células com cores (verde = acima da média, vermelho = abaixo)
-  Gera gráfico de barras embutido na planilha
-  Lê e escreve em Google Sheets via API (sem baixar o arquivo)
-  Exporta resumo em PDF
-  100% automático — roda com um único comando

---

## Estrutura

```
planilhas-bot/
├── excel_automation.py       # Automação completa de Excel
├── sheets_automation.py      # Leitura e escrita no Google Sheets
├── exemplos/
│   └── vendas_exemplo.xlsx   # Planilha de exemplo para testar
├── saida/                    # Relatórios gerados aqui
├── credentials.json          # Credenciais Google (você coloca aqui)
├── requirements.txt
└── README.md
```

---

## Como usar

### 1. Instale as dependências

```bash
pip install -r requirements.txt
```

### 2. Automação Excel

```bash
python excel_automation.py
```

Vai ler `exemplos/vendas_exemplo.xlsx` e gerar em `saida/`:
- `relatorio_formatado.xlsx` — planilha formatada com cores e gráfico
- `resumo.txt` — resumo em texto dos dados

### 3. Google Sheets

**Configurar credenciais Google:**
1. Acessa console.cloud.google.com
2. Cria um projeto → ativa **Google Sheets API** e **Google Drive API**
3. Cria uma **Service Account** → baixa o JSON → renomeia para `credentials.json`
4. Compartilha sua planilha com o e-mail da Service Account

```bash
python sheets_automation.py
```

---

## Adaptando para seu cliente

Edite no topo do `excel_automation.py`:

```python
ARQUIVO_ENTRADA = "sua_planilha.xlsx"   # Planilha do cliente
COLUNA_CATEGORIA = "Produto"            # Nome da coluna de categoria
COLUNA_VALOR = "Valor"                  # Nome da coluna de valores
COLUNA_DATA = "Data"                    # Nome da coluna de datas
```

---

## Tecnologias

| Tecnologia | Uso |
|---|---|
| Python 3.11+ | Linguagem principal |
| openpyxl | Leitura e escrita de Excel |
| pandas | Análise e processamento de dados |
| gspread | Integração com Google Sheets |
| google-auth | Autenticação Google |
| matplotlib | Geração de gráficos |

---

## Licença
MIT — livre para uso comercial.
