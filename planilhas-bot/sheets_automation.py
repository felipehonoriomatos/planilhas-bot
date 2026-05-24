import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import os

CREDENTIALS_FILE  = "credentials.json"      
SPREADSHEET_NAME  = "Planilha de Vendas"     
ABA_DADOS         = "Dados"                  
ABA_RELATORIO     = "Relatório Automático"   
COLUNA_CATEGORIA  = 1   
COLUNA_VALOR      = 4   

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def conectar_google_sheets():
    if not os.path.exists(CREDENTIALS_FILE):
        raise FileNotFoundError(
            f"Arquivo '{CREDENTIALS_FILE}' não encontrado!\n"
            "   Siga as instruções do README para gerar as credenciais."
        )

    creds  = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)
    print("Conectado ao Google Sheets")
    return client


def ler_dados(sheet) -> list[list]:
    dados = sheet.get_all_values()
    print(f"   → {len(dados) - 1} linhas encontradas (excluindo cabeçalho)")
    return dados


def processar_dados(dados: list[list]) -> dict:
    cabecalho = dados[0]
    linhas    = dados[1:]

    totais    = {}
    contagens = {}
    valores   = []

    for linha in linhas:
        try:
            categoria = linha[COLUNA_CATEGORIA - 1].strip()
            valor_str = linha[COLUNA_VALOR - 1].replace("R$", "").replace(",", ".").replace(" ", "")
            valor     = float(valor_str)

            totais[categoria]    = totais.get(categoria, 0) + valor
            contagens[categoria] = contagens.get(categoria, 0) + 1
            valores.append(valor)
        except (ValueError, IndexError):
            continue

    return {
        "totais":    totais,
        "contagens": contagens,
        "valores":   valores,
        "total_geral": sum(valores),
        "media":     sum(valores) / len(valores) if valores else 0,
        "maior":     max(valores) if valores else 0,
        "menor":     min(valores) if valores else 0,
    }


def escrever_relatorio(planilha, stats: dict):

    try:
        aba = planilha.worksheet(ABA_RELATORIO)
        aba.clear()
        print(f"   → Aba '{ABA_RELATORIO}' limpa e reutilizada")
    except gspread.WorksheetNotFound:
        aba = planilha.add_worksheet(title=ABA_RELATORIO, rows=100, cols=10)
        print(f"   → Aba '{ABA_RELATORIO}' criada")

    agora = datetime.now().strftime("%d/%m/%Y %H:%M")

    linhas = [
        [f"RELATÓRIO AUTOMÁTICO — Gerado em {agora}"],
        [],
        ["RESUMO GERAL"],
        ["Total de registros", len(stats["valores"])],
        ["Soma total",         f"R$ {stats['total_geral']:,.2f}"],
        ["Média por registro", f"R$ {stats['media']:,.2f}"],
        ["Maior valor",        f"R$ {stats['maior']:,.2f}"],
        ["Menor valor",        f"R$ {stats['menor']:,.2f}"],
        [],
        ["TOTAIS POR CATEGORIA"],
        ["Categoria", "Qtd Registros", "Total (R$)", "Média (R$)"],
    ]

    for cat, total in sorted(stats["totais"].items(), key=lambda x: -x[1]):
        qtd   = stats["contagens"][cat]
        media = total / qtd if qtd else 0
        linhas.append([
            cat,
            qtd,
            f"R$ {total:,.2f}",
            f"R$ {media:,.2f}",
        ])

    aba.update(range_name="A1", values=linhas)
    print(f"Relatório escrito na aba '{ABA_RELATORIO}'")


def main():
    print("\nAutomação Google Sheets iniciada...\n")

    try:
        client    = conectar_google_sheets()
        planilha  = client.open(SPREADSHEET_NAME)
        print(f"Planilha '{SPREADSHEET_NAME}' aberta")

        aba_dados = planilha.worksheet(ABA_DADOS)
        dados     = ler_dados(aba_dados)

        print("\nProcessando dados...")
        stats = processar_dados(dados)

        print(f"\nResultados:")
        print(f"   Total geral : R$ {stats['total_geral']:,.2f}")
        print(f"   Média       : R$ {stats['media']:,.2f}")
        print(f"   Categorias  : {list(stats['totais'].keys())}")

        print(f"\nEscrevendo relatório na planilha...")
        escrever_relatorio(planilha, stats)

        print("\nConcluído! Abra sua planilha Google para ver o relatório.")

    except FileNotFoundError as e:
        print(e)
    except gspread.SpreadsheetNotFound:
        print(f"Planilha '{SPREADSHEET_NAME}' não encontrada.")
        print("   Verifique o nome e se foi compartilhada com a Service Account.")
    except Exception as e:
        print(f"Erro inesperado: {e}")


if __name__ == "__main__":
    main()
