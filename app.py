import streamlit as st
import pandas as pd
from io import BytesIO

from main import processar_pdf
from diferencas import checar_divergencias
from openpyxl.styles import Font, PatternFill, Alignment, numbers
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Analisador de Ativos", layout="centered")
st.title("📄 Comparador de Ativos PDF x COMDINHEIRO")

pdf_file = st.file_uploader("📥 Envie o arquivo PDF do extrato", type="pdf")
excel_file = st.file_uploader("📊 Envie o Excel COMDINHEIRO", type=["xlsx"])

if st.button("🚀 Processar e Comparar") and pdf_file and excel_file:
    with st.spinner("Processando arquivos..."):
        try:
            # Processa o PDF
            df_ativos, excel_buffer = processar_pdf(pdf_file.read(), return_excel=True)
            st.success("✅ PDF processado com sucesso!")
            st.dataframe(df_ativos)

            # Processa o Excel COMDINHEIRO
            df_cd = pd.read_excel(excel_file)

            # Verifica divergências
            df_diferencas, report_buffer = checar_divergencias(df_ativos, df_cd)

            if df_diferencas.notna:
                st.success("✅ Relatório de Comparação gerado!")
                st.dataframe(df_diferencas)

                # Botão de download do relatório gerado
                st.download_button(
                    label="📥 Baixar Relatório Consolidado",
                    data=report_buffer,
                    file_name="relatorio_consolidado_equity.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Erro ao processar: {e}")
            st.error("Verifique os formatos dos arquivos e tente novamente.")
