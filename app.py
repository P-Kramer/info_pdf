import streamlit as st
import pandas as pd
from io import BytesIO
from main import processar_pdf
from diferencas import checar_divergencias
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from PIL import Image

st.set_page_config(page_title="Analisador de Ativos", layout="centered")

# ==== CABEÇALHO ====
logo_longview = Image.open("longview.png")
def mostrar_header():
    st.image(logo_longview, use_container_width=False, width=320)

mostrar_header()
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
    st.markdown("📄 Extrato em PDF (.pdf)")
    pdf_file = st.file_uploader("", type="pdf", key="pdf")

with col2:
    st.markdown("📊 Planilha COMDINHEIRO (.xlsx)")
    excel_file = st.file_uploader("", type=["xlsx"], key="excel")
    st.markdown("Colunas Necessárias: 'Carteira', 'Ativo', 'Descrição', 'Quant.', 'Saldo Bruto', 'Classe', 'ticker_cmd_puro'")

# ==== BOTÃO DE PROCESSAMENTO ====
st.markdown("---")
if st.button("🔍 Iniciar Comparação") and pdf_file and excel_file:
    with st.spinner("⏳ Processando arquivos..."):
        try:
            # 1) Extrai dados do PDF
            df_ativos, excel_buffer = processar_pdf(pdf_file.read(), return_excel=True)
            st.success("✅ PDF processado com sucesso!")

            with st.expander("📋 Visualizar dados extraídos do PDF"):
                st.dataframe(df_ativos, use_container_width=True)

            # 2) Lê Excel COMDINHEIRO
            df_cd = pd.read_excel(excel_file)

            # 3) Compara os dados
            df_diferencas, report_buffer = checar_divergencias(df_ativos, df_cd)

            # 4) Sempre tentar mostrar os PAREADOS a partir do relatório
            pareados_df = None
            try:
                report_buffer.seek(0)
                xls = pd.ExcelFile(report_buffer)
                if "Pareados" in xls.sheet_names:
                    pareados_df = pd.read_excel(xls, sheet_name="Pareados")
            except Exception as err:
                st.warning("⚠ Não foi possível ler a aba 'Pareados' do relatório gerado. Verifique o checar_divergencias().")
                pareados_df = None

            # 5) Mostrar divergências (se houver)
            if not df_diferencas.empty:
                st.success("✅ Comparação concluída com sucesso! Divergências encontradas.")
                with st.expander("🔎 Visualizar divergências encontradas"):
                    st.dataframe(df_diferencas, use_container_width=True)
            else:
                st.info("✅ Nenhuma divergência encontrada entre os dados.")

            # 6) Mostrar SEMPRE os pareados (se conseguimos ler)
            if pareados_df is not None and not pareados_df.empty:
                st.markdown("### ✅ Ativos pareados (sempre mostrado)")
                st.dataframe(pareados_df, use_container_width=True)
            else:
                st.warning("⚠ Relatório gerado não trouxe a aba 'Pareados' ou ela está vazia.")

            # 7) Botão de download SEMPRE, porque o cara pode querer ver as outras abas
            st.download_button(
                label="📥 Baixar Relatório em Excel",
                data=report_buffer.getvalue(),
                file_name="relatorio_consolidado_equity.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error("❌ Ocorreu um erro ao processar os arquivos.")
            st.exception(e)

# ==== RODAPÉ ====
st.markdown("---")
st.caption("Desenvolvido por Pedro Averame • Última atualização: Julho/2025")
