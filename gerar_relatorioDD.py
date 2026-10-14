import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

st.set_page_config(page_title="Gerador de Relatórios", layout="centered")

st.title("Gerador de Relatório PDF - Despesas Diretas")
st.write("Faça o upload da sua planilha Excel, selecione o contrato e gere o relatório em PDF (as despesas zeradas serão removidas).")

arquivo_excel = st.file_uploader("1. Carregue a sua planilha Excel (.xlsx)", type=["xlsx"])

if arquivo_excel is not None:
    try:
        df = pd.read_excel(arquivo_excel)
        
        colunas_base = ['ITEM', 'Conta contabil', 'HISTORICO', 'Total Geral', 'Total Geral.1']
        colunas_projetos = [col for col in df.columns if col not in colunas_base]
        
        st.write("---")
        st.write("### Configuração do Relatório")
        
        nome_contrato = st.selectbox("2. Selecione o Contrato / Projeto:", options=colunas_projetos)
        competencia = st.text_input("3. Competência / Período:", value="Setembro de 2026")
        
        if st.button("Gerar PDF"):
            df[nome_contrato] = pd.to_numeric(df[nome_contrato], errors='coerce').fillna(0)
            df_filtrado = df[df[nome_contrato] > 0].copy()
            
            if df_filtrado.empty:
                st.warning(f"O contrato '{nome_contrato}' não possui despesas maiores que zero.")
            else:
                colunas_relatorio = ['ITEM', 'Conta contabil', 'HISTORICO', 'Total Geral', nome_contrato]
                colunas_existentes = [col for col in colunas_relatorio if col in df_filtrado.columns]
                df_relatorio = df_filtrado[colunas_existentes].copy()

                df_relatorio[nome_contrato] = df_relatorio[nome_contrato].apply(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                if 'Total Geral' in df_relatorio.columns:
                    df_relatorio['Total Geral'] = df_relatorio['Total Geral'].apply(lambda x: f"{x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                if 'ITEM' in df_relatorio.columns:
                    df_relatorio['ITEM'] = df_relatorio['ITEM'].apply(lambda x: f"{int(x):03d}" if pd.notnull(x) else "")

                # Criação do PDF usando FPDF (sem depender de wkhtmltopdf)
                class PDF(FPDF):
                    def header(self):
                        self.set_font("helvetica", "B", 12)
                        self.cell(0, 6, "CEJAM", align="C", new_x="LMARGIN", new_y="NEXT")
                        self.set_font("helvetica", "B", 10)
                        self.cell(0, 6, "DEMONSTRATIVO FINANCEIRO DESPESA PORMENORIZADO", align="C", new_x="LMARGIN", new_y="NEXT")
                        self.cell(0, 6, f"DESPESAS DIRETAS - {nome_contrato}", align="C", new_x="LMARGIN", new_y="NEXT")
                        self.cell(0, 6, f"Período: {competencia}", align="C", new_x="LMARGIN", new_y="NEXT")
                        self.ln(5)

                # L = Landscape (Paisagem)
                pdf = PDF(orientation="L", unit="mm", format="A4")
                pdf.add_page()
                pdf.set_font("helvetica", size=7)
                
                # Cabeçalho da Tabela
                headers = df_relatorio.columns.tolist()
                
                # Calcular larguras das colunas (total A4 landscape tem ~277mm de largura útil)
                larguras = [10, 50, 150, 25, 42] # Ajuste fino das colunas
                # Se não houver 5 colunas, distribui por igual
                if len(headers) != 5:
                    larguras = [277/len(headers)] * len(headers)

                # Desenhar cabeçalho
                pdf.set_fill_color(230, 230, 230)
                pdf.set_font("helvetica", "B", 7)
                for i, heading in enumerate(headers):
                    pdf.cell(larguras[i], 6, str(heading)[:100], border=1, align="C", fill=True)
                pdf.ln()

                # Desenhar linhas de dados
                pdf.set_font("helvetica", size=7)
                for index, row in df_relatorio.iterrows():
                    # Para linhas com texto muito longo, podemos usar multi_cell se necessário, 
                    # mas para manter alinhado usaremos cell truncando ligeiramente ou ajustando o texto
                    for i, col_name in enumerate(headers):
                        texto = str(row[col_name])
                        # Truncar o histórico se for gigante para caber na linha (opcional)
                        if i == 2 and len(texto) > 110:
                            texto = texto[:107] + "..."
                        
                        align = "R" if i >= 3 else "L" # Números alinhados à direita
                        align = "C" if i == 0 else align # Item centrado
                        
                        pdf.cell(larguras[i], 5, texto, border=1, align=align)
                    pdf.ln()

                # Gerar ficheiro em memória (bytes)
                pdf_bytes = pdf.output(dest="S").encode("latin-1", "ignore")

                nome_limpo = nome_contrato.replace(' - ', '_').replace(' ', '_').replace('/', '-')
                nome_arquivo_pdf = f"Relatorio_{nome_limpo}_{competencia.replace(' ', '_')}.pdf"

                st.success("PDF gerado com sucesso!")
                
                st.download_button(
                    label="📥 Descarregar Relatório em PDF",
                    data=pdf_bytes,
                    file_name=nome_arquivo_pdf,
                    mime="application/pdf"
                )
                
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar a planilha: {e}")
