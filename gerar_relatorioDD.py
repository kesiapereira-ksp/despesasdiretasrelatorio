import streamlit as st
import pandas as pd
from fpdf import FPDF
from fpdf.fonts import FontFace
import os

st.set_page_config(page_title="Gerador de Relatórios", layout="centered")

st.title("Gerador de Relatório PDF - Despesas Diretas")
st.write("Faça o upload da sua planilha Excel, selecione um ou mais contratos e gere o relatório em PDF timbrado.")

arquivo_excel = st.file_uploader("1. Carregue a sua planilha Excel (.xlsx)", type=["xlsx"])

if arquivo_excel is not None:
    try:
        # Ler a planilha e limpar linhas vazias ou de totais antigos
        df = pd.read_excel(arquivo_excel)
        df = df.dropna(subset=['HISTORICO'])
        df = df[~df['HISTORICO'].astype(str).str.contains('Total Geral', case=False, na=False)]

        colunas_base = ['ITEM', 'Conta contabil', 'HISTORICO', 'Total Geral', 'Total Geral.1']
        colunas_projetos = [col for col in df.columns if col not in colunas_base]
        
        st.write("---")
        st.write("### Configuração do Relatório")
        
        projetos_selecionados = st.multiselect("2. Selecione os Contratos / Projetos:", options=colunas_projetos)
        competencia = st.text_input("3. Competência / Período:", value="Setembro de 2026")
        
        if st.button("Gerar PDF"):
            if not projetos_selecionados:
                st.warning("Por favor, selecione pelo menos um projeto.")
            else:
                class PDF(FPDF):
                    def header(self):
                        if os.path.exists("cabecalho.png"):
                            self.image("cabecalho.png", x=10, y=8, w=277)
                            self.ln(35) # Espaço do cabeçalho
                        else:
                            self.set_font("helvetica", "B", 12)
                            self.cell(0, 6, "CEJAM", align="C", new_x="LMARGIN", new_y="NEXT")
                            self.ln(10)

                    def footer(self):
                        self.set_y(-25)
                        if os.path.exists("rodape.png"):
                            self.image("rodape.png", x=10, y=self.get_y(), w=277)
                        else:
                            self.set_font("helvetica", "I", 8)
                            self.cell(0, 10, f"Página {self.page_no()}", align="C")

                pdf = PDF(orientation="L", unit="mm", format="A4")
                
                for nome_contrato in projetos_selecionados:
                    df_projeto = df.copy()
                    df_projeto[nome_contrato] = pd.to_numeric(df_projeto[nome_contrato], errors='coerce').fillna(0)
                    df_projeto['Total Geral'] = pd.to_numeric(df_projeto['Total Geral'], errors='coerce').fillna(0)
                    
                    df_filtrado = df_projeto[df_projeto[nome_contrato] > 0].copy()
                    
                    if not df_filtrado.empty:
                        pdf.add_page()
                        
                        pdf.set_font("helvetica", "B", 10)
                        pdf.cell(0, 6, "DEMONSTRATIVO FINANCEIRO DESPESA PORMENORIZADO", align="C", new_x="LMARGIN", new_y="NEXT")
                        pdf.cell(0, 6, f"DESPESAS DIRETAS - {nome_contrato}", align="C", new_x="LMARGIN", new_y="NEXT")
                        pdf.cell(0, 6, f"Período: {competencia}", align="C", new_x="LMARGIN", new_y="NEXT")
                        pdf.ln(5)

                        colunas_relatorio = ['Conta contabil', 'HISTORICO', 'Total Geral', nome_contrato]
                        colunas_existentes = [col for col in colunas_relatorio if col in df_filtrado.columns]
                        
                        total_geral_calc = df_filtrado['Total Geral'].sum()
                        total_projeto_calc = df_filtrado[nome_contrato].sum()
                        
                        nova_linha_total = pd.DataFrame({
                            'Conta contabil': [''],
                            'HISTORICO': ['TOTAL GERAL'],
                            'Total Geral': [total_geral_calc],
                            nome_contrato: [total_projeto_calc]
                        })
                        df_filtrado = pd.concat([df_filtrado, nova_linha_total], ignore_index=True)

                        def formata_moeda(valor):
                            if pd.isna(valor) or valor == '': return ''
                            if isinstance(valor, str): return valor
                            return f"{valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')

                        df_filtrado['Total Geral'] = df_filtrado['Total Geral'].apply(formata_moeda)
                        df_filtrado[nome_contrato] = df_filtrado[nome_contrato].apply(formata_moeda)

                        pdf.set_font("helvetica", size=7)
                        larguras = (50, 150, 35, 42)
                        
                        # Criar objeto FontFace para o negrito nas tabelas (Resolve o erro)
                        estilo_negrito = FontFace(emphasis="B")
                        
                        with pdf.table(col_widths=larguras, text_align=("LEFT", "LEFT", "RIGHT", "RIGHT")) as table:
                            # Aplicar negrito na linha do cabeçalho
                            cabecalho = table.row(style=estilo_negrito)
                            for col in colunas_existentes:
                                cabecalho.cell(str(col))
                            
                            for index, row in df_filtrado.iterrows():
                                is_total_row = (row['HISTORICO'] == 'TOTAL GERAL')
                                # Se for o "TOTAL GERAL", aplica o negrito na linha inteira
                                if is_total_row:
                                    linha_tabela = table.row(style=estilo_negrito)
                                else:
                                    linha_tabela = table.row()
                                    
                                for col in colunas_existentes:
                                    texto = str(row[col]) if pd.notnull(row[col]) else ""
                                    linha_tabela.cell(texto)

                pdf_bytes = bytes(pdf.output())

                nome_arquivo_pdf = f"Relatorio_Despesas_Diretas_{competencia.replace(' ', '_')}.pdf"

                st.success("PDF gerado com sucesso!")
                
                st.download_button(
                    label="📥 Descarregar Relatório em PDF",
                    data=pdf_bytes,
                    file_name=nome_arquivo_pdf,
                    mime="application/pdf"
                )
                
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar a planilha: {e}")
