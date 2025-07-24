import streamlit as st
import pandas as pd
from io import BytesIO
from main import processar_pdf
from diferencas import checar_divergencias
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Analisador de Ativos", layout="centered")

# ==== CABEÇALHO ====
st.markdown("## 🧾 Comparador de Ativos: PDF vs COMDINHEIRO")
st.markdown(
    """
    Esta ferramenta compara os ativos de um extrato em PDF com os dados do sistema COMDINHEIRO, 
    identificando divergências de valor, quantidade ou identificação.
    """
)

# ==== UPLOADS ====
st.markdown("### 📁 Upload dos Arquivos")

col1, col2 = st.columns(2)

with col1:
    st.markdown ("📄 Extrato em PDF (.pdf)")
    pdf_file = st.file_uploader("", type="pdf", key="pdf")

with col2:
    st.markdown ("📊 Planilha COMDINHEIRO (.xlsx)")
    excel_file = st.file_uploader("", type=["xlsx"], key="excel")
    st.markdown("Colunas Necessárias: 'Carteira', 'Ativo', 'Descrição', 'Quant.', 'Saldo Bruto', 'Classe', 'ticker_cmd_puro'")
# ==== BOTÃO DE PROCESSAMENTO ====
st.markdown("---")
if st.button("🔍 Iniciar Comparação") and pdf_file and excel_file:
    with st.spinner("⏳ Processando arquivos..."):
        try:
            # Extrai dados do PDF
            df_ativos, excel_buffer = processar_pdf(pdf_file.read(), return_excel=True)
            st.success("✅ PDF processado com sucesso!")

            with st.expander("📋 Visualizar dados extraídos do PDF"):
                st.dataframe(df_ativos, use_container_width=True)

            # Lê Excel COMDINHEIRO
            df_cd = pd.read_excel(excel_file)

            # Compara os dados
            df_diferencas, report_buffer = checar_divergencias(df_ativos, df_cd)

            if not df_diferencas.empty:
                st.success("✅ Comparação concluída com sucesso!")
                
                with st.expander("🔎 Visualizar divergências encontradas"):
                    st.dataframe(df_diferencas, use_container_width=True)

                st.download_button(
                    label="📥 Baixar Relatório em Excel",
                    data=report_buffer,
                    file_name="relatorio_consolidado_equity.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            else:
                st.info("✅ Nenhuma divergência encontrada entre os dados.")

        except Exception as e:
            st.error("❌ Ocorreu um erro ao processar os arquivos.")
            st.exception(e)

# ==== RODAPÉ ====
st.markdown("---")
st.caption("Desenvolvido por Pedro Averame • Última atualização: Julho/2025")
