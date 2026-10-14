import streamlit as st
import pandas as pd
import pdfkit

# Configuração da página web
st.set_page_config(page_title="Gerador de Relatórios", layout="centered")

st.title("Gerador de Relatório PDF - Despesas Diretas")
st.write("Faça o upload da sua planilha Excel para gerar o relatório em PDF com as despesas zeradas removidas.")

# 1. Campos para o utilizador preencher (pode alterar no ecrã antes de gerar)
nome_contrato = st.text_input("Nome da Coluna do Contrato (exato como está no Excel)", value="Cejam - Apoio à Gestão")
competencia = st.text_input("Competência / Período", value="Setembro de 2026")

# 2. Botão para carregar o ficheiro Excel
arquivo_excel = st.file_uploader("Carregue a sua planilha Excel (.xlsx)", type=["xlsx"])

# Só executa se o utilizador tiver carregado um ficheiro
if arquivo_excel is not None:
    if st.button("Processar e Gerar PDF"):
        try:
            # Ler a planilha diretamente do ficheiro carregado
            df = pd.read_excel(arquivo_excel)
            
            # Verificar se a coluna digitada realmente existe na planilha
            if nome_contrato not in df.columns:
                st.error(f"A coluna '{nome_contrato}' não foi encontrada na planilha. Verifique os espaços ou letras maiúsculas/minúsculas.")
            else:
                # 3. Garantir que a coluna é numérica e filtrar os valores zerados
                df[nome_contrato] = pd.to_numeric(df[nome_contrato], errors='coerce').fillna(0)
                df_filtrado = df[df[nome_contrato] > 0].copy()
                
                # 4. Formatar os valores para o padrão brasileiro (R$)
                df_filtrado[nome_contrato] = df_filtrado[nome_contrato].apply(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                if 'Total Geral' in df_filtrado.columns:
                    df_filtrado['Total Geral'] = df_filtrado['Total Geral'].apply(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

                # 5. Criar a estrutura HTML
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
                    <h3>Período {competencia}</h3>
                    {html_tabela}
                </body>
                </html>
                """

                # 6. Configurar o PDF em modo Paisagem
                opcoes = {
                    'page-size': 'A4',
                    'orientation': 'Landscape',
                    'encoding': "UTF-8",
                }
                
                # Gerar o PDF na memória (passando 'False' em vez de um nome de ficheiro para guardar)
                pdf_bytes = pdfkit.from_string(html_conteudo, False, options=opcoes)
                
                # Nome que o ficheiro terá ao ser descarregado
                nome_arquivo_pdf = f"Relatorio_{nome_contrato.replace(' - ', '_').replace(' ', '_')}_{competencia.replace(' ', '_')}.pdf"

                st.success("PDF gerado com sucesso!")
                
                # 7. Criar o botão de download para o utilizador descarregar o ficheiro
                st.download_button(
                    label="📥 Descarregar Relatório em PDF",
                    data=pdf_bytes,
                    file_name=nome_arquivo_pdf,
                    mime="application/pdf"
                )
                
        except Exception as e:
            st.error(f"Ocorreu um erro ao processar o ficheiro: {e}")
