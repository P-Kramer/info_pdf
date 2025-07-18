import fitz  # PyMuPDF
import pandas as pd
import re
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Caminho do PDF
pdf_path = "C:/Users/Pedro Averame/Documents/STATEMENT_TEST.pdf"
dados = []

# Função para verificar se o span está em negrito
def is_bold(span):
    return "Bold" in span.get("font", "")

ativo_atual = None
total_extraido = False
linhas_ativas = []

with fitz.open(pdf_path) as doc:
    for i, page in enumerate(doc):
        spans = []
        blocks = page.get_text("dict")["blocks"]
        for block in blocks:
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    spans.append(span)

        idx = 0
        while idx < len(spans):
            span = spans[idx]
            texto = span["text"].strip()

            # Detecta novo ativo válido
            if is_bold(span) and "(" in texto and ")" in texto:
                if any(palavra in texto for palavra in ["$", "/", "(MMF)", "(NL)", ",", "Approved List", "Focus List", "Investment Objectives", "Asset Class"]):
                    idx += 1
                    continue

                # Se havia ativo anterior sem total, tenta pegar última linha válida com números
                if ativo_atual and not total_extraido and linhas_ativas:
                    c= 0
                    for linha in (linhas_ativas):
                        textos = [span.get("text", "").strip() for span in linha if "text" in span]

                        while c == 0:
                            infos = textos
                            c+=1
                            try:
                                
                                quantidade = float(infos[1].replace(",", ""))
                                total_cost = float(infos[4].replace(",", ""))
                                market_value = float(infos[5].replace(",", ""))
                                tickers = re.findall(r"\(([^)]+)\)", ativo_atual)
                                ticker = tickers[-1].strip() if tickers else ""

                                dados.append({
                                    "Página": i + 1,
                                    "Ativo": ativo_atual,
                                    "Ticker": ticker,
                                    "Quantidade Total": quantidade,
                                    "Total Cost": total_cost,
                                    "Market Value": market_value
                                })
                                break
                            except Exception as e:
                                print(f"Erro na extração alternativa de {ativo_atual}: {e}")

                # Atualiza o novo ativo
                ativo_atual = texto
                total_extraido = False
                linhas_ativas = []
                idx += 1
                continue

            # Coleta linha se parecer conter números
            if ativo_atual and re.search(r"\d", texto):
                linha_nova = spans[idx:idx+6]
                linhas_ativas.append(linha_nova)

            # Detecta linha de "Total"
            if ativo_atual and texto.startswith("Total"):
                texto_total = ' '.join([sp["text"] for sp in spans[idx:idx+6]])
                numeros = re.findall(r"[-]?[\d,.]+", texto_total)
                if len(numeros) >= 3:
                    try:
                        quantidade = float(numeros[0].replace(",", ""))
                        total_cost = float(numeros[1].replace(",", ""))
                        market_value = float(numeros[2].replace(",", ""))
                        tickers = re.findall(r"\(([^)]+)\)", ativo_atual)
                        ticker = tickers[-1].strip() if tickers else ""

                        dados.append({
                            "Página": i + 1,
                            "Ativo": ativo_atual,
                            "Ticker": ticker,
                            "Quantidade Total": quantidade,
                            "Total Cost": total_cost,
                            "Market Value": market_value
                        })

                        total_extraido = True
                        ativo_atual = None
                        linhas_ativas = []

                    except Exception as e:
                        print(f"Erro ao processar '{ativo_atual}': {e}")
            idx += 1

# Cria e exporta o DataFrame
df = pd.DataFrame(dados)
excel_path = "ativos_extraidos_corrigido_final.xlsx"
with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Ativos")
    workbook = writer.book
    worksheet = writer.sheets["Ativos"]

    border = Border(left=Side(style='thin', color='000000'),
                    right=Side(style='thin', color='000000'),
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='thin', color='000000'))

    header_font = Font(bold=True, color="000000")
    header_fill = PatternFill("solid", fgColor="FF3D98FF")  # FF = opaco


    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = border
            if cell.row == 1:
                cell.font = header_font
                cell.fill = header_fill

    for col in worksheet.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        worksheet.column_dimensions[col_letter].width = max_length + 2

print(f"✅ Extração e formatação concluídas. Arquivo gerado: {excel_path}")
