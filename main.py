import streamlit as st
import fitz  # PyMuPDF
import pandas as pd

# Substitua pelo caminho do seu PDF local
pdf_path = "C:\\Users\\Pedro Averame\\Documents\\STATEMENT_TEST.pdf"


# Lista para armazenar os resultados
dados = []

# Função auxiliar para extrair dados de cada página
def extrair_dados(texto_pagina):
    linhas = texto_pagina.split('\n')
    ativo = None

    for i, linha in enumerate(linhas):
        # Detecta nome do ativo
        if "Next Dividend Payable" in linha and "Asset Class" in linha:
            proxima = linhas[i + 1] if i + 1 < len(linhas) else ""
            if "(" in proxima and ")" in proxima:
                ativo = proxima.strip()
        # Detecta linha de total
        if linha.strip().startswith("Total"):
            partes = linha.strip().split()
            try:
                quantidade_total = float(partes[1].replace(',', ''))
                total_cost = float(partes[2].replace(',', ''))
                market_value = float(partes[3].replace(',', ''))

                if ativo:
                    ticker = ativo.split('(')[-1].strip(')')
                    dados.append({
                        "Ativo": ativo,
                        "Ticker": ticker,
                        "Quantidade Total": quantidade_total,
                        "Total Cost": total_cost,
                        "Market Value": market_value
                    })
            except:
                continue

# Abrir o PDF
with fitz.open(pdf_path) as doc:
    for pagina in doc:
        texto = pagina.get_text()
        extrair_dados(texto)

# Converter em DataFrame e salvar como Excel
df = pd.DataFrame(dados)
df.to_excel("ativos_extraidos.xlsx", index=False)

print("Extração concluída. Arquivo salvo como 'ativos_extraidos.xlsx'")