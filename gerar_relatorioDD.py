import streamlit as st
import pandas as pd
from fpdf import FPDF
from fpdf.fonts import FontFace
import io

st.set_page_config(page_title="Gerador de Relatórios", layout="centered")

st.title("Gerador de Relatório PDF - Despesas Diretas")
st.write("Faça o upload da sua planilha Excel, anexe o papel timbrado, selecione os contratos e gere os relatórios em PDFs separados (Vertical).")

st.write("### 1. Arquivos Base")
arquivo_excel = st.file_uploader("Carregue a sua planilha Excel (.xlsx)", type=["xlsx"])

# Criar duas colunas para colocar os botões de upload das imagens lado a lado
col1, col2 = st.columns(2)
with col1:
    img_cabecalho = st.file_uploader("Imagem do Cabeçalho (Opcional)", type=["png", "jpg", "jpeg"])
with col2:
    img_rodape = st.file_uploader("Imagem do Rodapé (Opcional)", type=["png", "jpg", "jpeg"])

if arquivo_excel is not None:
    try:
        # Ler a planilha e limpar linhas vazias ou de totais antigos
        df = pd.read_excel(arquivo_excel)
        df = df.dropna(subset=['HISTORICO'])
        df = df[~df['HISTORICO'].astype(str).str.contains('Total Geral', case=False, na=False)]

        colunas_base = ['ITEM', 'Conta contabil', 'HISTORICO', 'Total Geral', 'Total Geral.1']
        colunas_projetos = [col for col in df.columns if col not in colunas_base]
        
        st.write("---")
        st.write("### 2. Configuração do Relatório")
        
        projetos_selecionados = st.multiselect("Selecione os Contratos / Projetos:", options=colunas_projetos)
        competencia = st.text_input("Competência / Período:", value="Setembro de 2026")
        
        if st.button("Gerar PDFs"):
            if not projetos_selecionados:
                st.warning("Por favor, selecione pelo menos um projeto.")
            else:
                # Prepara as imagens na memória de forma segura para o PDF
                cabecalho_stream = io.BytesIO(img_cabecalho.read()) if img_cabecalho else None
                rodape_stream = io.BytesIO(img_rodape.read()) if img_rodape else None

                pdfs_gerados = [] # Lista para guardar os relatórios gerados

                for nome_contrato in projetos_selecionados:
                    df_projeto = df.copy()
                    df_projeto[nome_contrato] = pd.to_numeric(df_projeto[nome_contrato], errors='coerce').fillna(0)
                    df_projeto['Total Geral'] = pd.to_numeric(df_projeto['Total Geral'], errors='coerce').fillna(0)
                    
                    # Filtra apenas despesas > 0 para este contrato específico
                    df_filtrado = df_projeto[df_projeto[nome_contrato] > 0].copy()
                    
                    if not df_filtrado.empty:
                        # Criar a classe PDF dentro do loop para garantir que é limpa a cada ficheiro gerado
                        class PDF(FPDF):
                            def header(self):
                                if cabecalho_stream is not None:
                                    # Usa w=190 (largura útil da página A4 vertical)
                                    cabecalho_stream.seek(0)
                                    self.image(cabecalho_stream, x=10, y=8, w=190)
                                    self.ln(35) # Ajuste este valor (35) se o seu cabeçalho for mais alto ou mais baixo
                                else:
                                    self.set_font("helvetica", "B", 12)
                                    self.cell(0, 6, "CEJAM", align="C", new_x="LMARGIN", new_y="NEXT")
                                    self.ln(10)

                            def footer(self):
                                self.set_y(-25)
                                if rodape_stream is not None:
                                    rodape_stream.seek(0)
                                    self.image(rodape_stream, x=10, y=self.get_y(), w=190)
                                else:
                                    self.set_font("helvetica", "I", 8)
                                    self.cell(0, 10, f"Página {self.page_no()}", align="C")

                        # "P" = Portrait (Vertical)
                        pdf = PDF(orientation="P", unit="mm", format="A4")
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
                        
                        # Ajuste das larguras para formato Vertical (Total máximo = 190mm)
                        # Conta(30) + Historico(100) + Total Geral(30) + Contrato(30) = 190mm
                        larguras = (30, 100, 30, 30)
                        
                        estilo_negrito = FontFace(emphasis="B")
                        
                        with pdf.table(col_widths=larguras, text_align=("LEFT", "LEFT", "RIGHT", "RIGHT")) as table:
                            cabecalho = table.row(style=estilo_negrito)
                            for col in colunas_existentes:
                                cabecalho.cell(str(col))
                            
                            for index, row in df_filtrado.iterrows():
                                is_total_row = (row['HISTORICO'] == 'TOTAL GERAL')
                                if is_total_row:
                                    linha_tabela = table.row(style=estilo_negrito)
                                else:
                                    linha_tabela = table.row()
                                    
                                for col in colunas_existentes:
                                    texto = str(row[col]) if pd.notnull(row[col]) else ""
                                    linha_tabela.cell(texto)

                        pdf_bytes = bytes(pdf.output())

                        nome_limpo = nome_contrato.replace(' - ', '_').replace(' ', '_').replace('/', '-')
                        nome_arquivo_pdf = f"Relatorio_Diretas_{nome_limpo}_{competencia.replace(' ', '_')}.pdf"
                        
                        # Guarda na lista de ficheiros gerados
                        pdfs_gerados.append((nome_arquivo_pdf, pdf_bytes))

                # Apresentar botões de download separados
                if not pdfs_gerados:
                    st.warning("Nenhum PDF foi gerado (os projetos selecionados não tinham despesas ou o valor estava zerado).")
                else:
                    st.success(f"{len(pdfs_gerados)} PDF(s) gerado(s) com sucesso!")
                    
                    for nome_arq, byts in pdfs_gerados:
                        st.download_button(
                            label=f"📥 Descarregar {nome_arq}",
                            data=byts,
                            file_name=nome_arq,
                            mime="application/pdf",
                            key=nome_arq # Identificador único para cada botão na tela
                        )
                
    except Exception as e:
        st.error(f"Ocorreu um erro ao processar a planilha: {e}")
