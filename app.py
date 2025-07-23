import streamlit as st
import pandas as pd
from io import BytesIO

from main import processar_pdf
from diferencas import checar_divergencias

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
            resultado = checar_divergencias(df_ativos, df_cd)
            
            if resultado is None or not isinstance(resultado, tuple) or len(resultado) != 2:
                raise ValueError("A função `checar_divergencias` não retornou os dois valores esperados.")

            df_diferencas, report_buffer = resultado

            if df_diferencas.empty:
                st.success("🎉 Nenhuma diferença encontrada!")

            else:
                st.warning("⚠️ Diferenças encontradas:")
                st.dataframe(df_diferencas.head())

                # Botões para download
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📥 Baixar Relatório de Diferenças",
                        data=report_buffer,
                        file_name="relatorio_diferencas.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Erro ao processar: {e}")
            st.error("Verifique os formatos dos arquivos e tente novamente.")
