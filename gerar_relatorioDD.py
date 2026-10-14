import pandas as pd
import pdfkit

# Variáveis que definem o nome do arquivo final e a coluna lida
ARQUIVO_EXCEL = 'dados/planilha.xlsx'
NOME_CONTRATO = 'Cejam - Apoio à Gestão'
COMPETENCIA = 'Setembro de 2026'

def processar_relatorio():
    # 1. Carregar a planilha
    df = pd.read_excel(ARQUIVO_EXCEL)
    
    # 2. Garantir que a coluna do contrato é numérica e filtrar os valores zerados
    df[NOME_CONTRATO] = pd.to_numeric(df[NOME_CONTRATO], errors='coerce').fillna(0)
    df_filtrado = df[df[NOME_CONTRATO] > 0].copy()
    
    # 3. Formatar os valores para o padrão brasileiro (R$)
    df_filtrado[NOME_CONTRATO] = df_filtrado[NOME_CONTRATO].apply(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    if 'Total Geral' in df_filtrado.columns:
        df_filtrado['Total Geral'] = df_filtrado['Total Geral'].apply(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

    # 4. Criar a estrutura visual do PDF baseada na sua imagem
    html_tabela = df_filtrado.to_html(index=False, justify='center')
    html_conteudo = f"""
    <html>
    <head>
        <meta charset="UTF-8">
        <style>
            body {{ font-family: Arial, sans-serif; font-size: 10px; }}
            h2, h3 {{ text-align: center; margin: 5px 0; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid black; padding: 5px; text-align: left; }}
            th {{ background-color: #f2f2f2; }}
        </style>
    </head>
    <body>
        <h2>CEJAM</h2>
        <h3>DEMONSTRATIVO FINANCEIRO DESPESA PORMENORIZADO</h3>
        <h3>DESPESAS DIRETAS</h3>
        <h3>Período {COMPETENCIA}</h3>
        {html_tabela}
    </body>
    </html>
    """

    # 5. Gerar o nome do arquivo com Contrato e Competência
    nome_pdf = f"{NOME_CONTRATO.replace(' - ', '_').replace(' ', '_')}_{COMPETENCIA.replace(' ', '_')}.pdf"
    
    # 6. Configurar e salvar o PDF em modo Paisagem
    opcoes = {
        'page-size': 'A4',
        'orientation': 'Landscape',
        'encoding': "UTF-8",
    }
    
    pdfkit.from_string(html_conteudo, nome_pdf, options=opcoes)

if __name__ == "__main__":
    processar_relatorio()
