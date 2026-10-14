import streamlit as st
import pandas as pd
import pdfkit

# Configuração da página web
st.set_page_config(page_title="Gerador de Relatórios", layout="centered")

st.title("Gerador de Relatório PDF - Despesas DIretas")
st.write("Faça o upload da sua planilha Excel, selecione o contrato através das colunas e gere o relatório em PDF (as despesas zeradas serão removidas).")

# 1. Carregar o ficheiro Excel
arquivo_excel = st.file_uploader("1. Carregue a sua planilha Excel (.xlsx)", type=["xlsx"])

if arquivo_excel is not None:
    try:
        # Ler a planilha
        df = pd.read_excel(arquivo_excel)
        
        # Identificar as colunas que representam os projetos/contratos
        # Excluímos as colunas base para que apenas os nomes dos projetos apareçam nas opções
        colunas_base = ['ITEM', 'Conta contabil', 'HISTORICO', 'Total Geral', 'Total Geral.1']
        colunas_projetos = [col for col in df.columns if col not in colunas_base]
        
        st.write("---")
        st.write("### Configuração do Relatório")
        
        # 2. Dropdown para o utilizador escolher o projeto com base nas colunas do próprio Excel carregado
        nome_contrato = st.selectbox("2. Selecione o Contrato / Projeto:", options=colunas_projetos)
        
        # 3. Campo de Competência
        competencia = st.text_input("3. Competência / Período:", value="Setembro de 2026")
        
        if st.button("Gerar PDF"):
            # Garantir que a coluna escolhida é numérica e filtrar apenas os valores > 0
            df[nome_contrato] = pd.to_numeric(df[nome_contrato], errors='coerce').fillna(0)
            df_filtrado = df[df[nome_contrato] > 0].copy()
            
            if df_filtrado.empty:
                st.warning(f"O contrato '{nome_contrato}' não possui despesas maiores que zero.")
            else:
                # Filtrar apenas as colunas essenciais para o relatório em PDF
                colunas_relatorio = ['ITEM', 'Conta contabil', 'HISTORICO', 'Total Geral', nome_contrato]
                
                # Tentar garantir que todas as colunas existem (caso falte o ITEM ou HISTORICO, não quebra o código)
                colunas_existentes = [col for col in colunas_relatorio if col in df_filtrado.columns]
                df_relatorio = df_filtrado[colunas_existentes].copy()

                # Formatação visual (Dinheiro com R$ e remover casas decimais do ITEM)
                df_relatorio[nome_contrato] = df_relatorio[nome_contrato].apply(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                if 'Total Geral' in df_relatorio.columns:
                    df_relatorio['Total Geral'] = df_relatorio['Total Geral'].apply(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                if 'ITEM' in df_relatorio.columns:
                    df_relatorio['ITEM'] = df_relatorio['ITEM'].apply(lambda x: f"{int(x):03d}" if pd.notnull(x) else "")

                # 4. Criar a estrutura HTML
                html_tabela = df_relatorio.to_html(index=False, justify='center')
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
                    <h3>DESPESAS DIRETAS - {nome_contrato}</h3>
                    <h3>Período: {competencia}</h3>
                    {html_tabela}
                </body>
                </html>
                """

                # 5. Configurar o PDF em modo Paisagem
                opcoes = {
                    'page-size': 'A4',
                    'orientation': 'Landscape',
                    'encoding': "UTF-8",
                }
                
                # Gerar o PDF na memória
                pdf_bytes = pdfkit.from_string(html_conteudo, False, options=opcoes)
                
                # Nome do ficheiro final a ser descarregado (Limpa caracteres inválidos como as barras de datas)
                nome_limpo = nome_contrato.replace(' - ', '_').replace(' ', '_').replace('/', '-')
                nome_arquivo_pdf = f"Relatorio_{nome_limpo}_{competencia.replace(' ', '_')}.pdf"

                st.success("PDF gerado com sucesso! Clique no botão abaixo para o descarregar.")
                
                # 6. Criar o botão de download
                st.download_button(
                    label="📥 Descarregar Relatório em PDF",
                    data=pdf_bytes,
                    file_name=nome_arquivo_pdf,
                    mime="application/pdf"
                )
                
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar a planilha: {e}")
