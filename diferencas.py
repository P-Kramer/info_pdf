
import pandas as pd
import numpy as np
from rapidfuzz import fuzz, process
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Configurações
RED_FILL = PatternFill(start_color='FFFFF000', end_color='FFFFF000', fill_type='solid')

# Mapeamento manual de pareamentos forçados
pareamentos_forcados = {
    "AMERCO /NV/":"UHAL'B",
    "POLEN CAPITAL FOCUS US GR USD INSTL":"PFUGI",
    "JPM US SELECT EQUITY PLUS C (ACC) USD":"SELQZ",
    "AMUNDI FDS US EQ FUNDM GR I2 USD C":"PONVS",
    "MS INVF GLOBAL ENDURANCE I USD ACC":"SMFVZ"
}

def processar_com_dinheiro(df):
    df = df[
        (df['Carteira'].str.startswith('LA_', na=False)) &
        (df['Classe'].isin(['EQUITY', 'ALTERNATIVES', "FIXED INCOME","FLOATING INCOME"]))
    ].copy()
    df = df.rename(columns={
        'Ativo': 'Ticker',
        'Saldo Bruto (quant) em 30/06/2025': 'Quantidade',
        'Saldo Bruto em 30/06/2025': 'MarketValue'
    })
    df['TickerBase'] = df['ticker_cmd_puro'].str.split(':').str[-1].str.strip()
    return df[['Descrição', 'Ticker', 'TickerBase', 'Quantidade', 'MarketValue', 'Classe']]

def processar_ativos(df):
    df = df.rename(columns={
        'Ativo': 'Nome',
        'Quantidade Total': 'Quantidade',
        'Market Value': 'MarketValue'
    })
    df['TickerBase'] = df['Ticker'].str.extract(r'([A-Z]{2,6}$)')[0].fillna(df['Ticker'])
    return df[['Nome', 'Ticker', 'TickerBase', 'Quantidade', 'MarketValue']]

ativos_path = 'ativos_extraidos_corrigido_final.xlsx'
cd_path = 'COMDINHEIRO.xlsx'
df_cd = pd.read_excel(cd_path, sheet_name='COM DINHEIRO')
df_at = pd.read_excel(ativos_path, sheet_name='Ativos')

equity_cd = processar_com_dinheiro(df_cd)
equity_at = processar_ativos(df_at)

# Matching exato
exact_match = pd.merge(
    equity_cd,
    equity_at,
    on='TickerBase',
    suffixes=('_CD', '_MS'),
    how='inner'
)
matched_tickers = exact_match['TickerBase'].unique()

remaining_cd = equity_cd[~equity_cd['TickerBase'].isin(matched_tickers)]
remaining_at = equity_at[~equity_at['TickerBase'].isin(matched_tickers)]

# Matching forçado
forcados = []
for _, row in remaining_cd.iterrows():
    descricao_upper = row['Descrição'].strip().upper()
    if descricao_upper in pareamentos_forcados:
        ticker_alvo = pareamentos_forcados[descricao_upper]
        match_row = remaining_at[remaining_at['Ticker'] == ticker_alvo]
        if not match_row.empty:
            match_row = match_row.iloc[0]
            forcados.append({
                'Descrição_CD': row['Descrição'],
                'Ticker_CD': row['Ticker'],
                'TickerBase': row['TickerBase'],
                'Classe': row ["Classe"],
                'Quantidade_CD': row['Quantidade'],
                'MarketValue_CD': row['MarketValue'],
                'Nome_MS': match_row['Nome'],
                'Ticker_MS': match_row['Ticker'],
                'Quantidade_MS': match_row['Quantidade'],
                'MarketValue_MS': match_row['MarketValue'],
                'Similaridade': 100
            })

# Fuzzy match por descrição completa
fuzzy_matches = []
for _, row in remaining_cd.iterrows():
    match, score, _ = process.extractOne(
        row['Descrição'],
        remaining_at['Nome'].tolist(),
        scorer=fuzz.token_set_ratio
    )
    if score >= 85:
        match_row = remaining_at[remaining_at['Nome'] == match].iloc[0]
        fuzzy_matches.append({
            'Descrição_CD': row['Descrição'],
            'Ticker_CD': row['Ticker'],
            'TickerBase': row['TickerBase'],
            'Classe': row ["Classe"],
            'Quantidade_CD': row['Quantidade'],
            'MarketValue_CD': row['MarketValue'],
            'Nome_MS': match_row['Nome'],
            'Ticker_MS': match_row['Ticker'],
            'Quantidade_MS': match_row['Quantidade'],
            'MarketValue_MS': match_row['MarketValue'],
            'Similaridade': score
        })

# Junta exato + forçado + fuzzy descrição
fuzzy_df = pd.DataFrame(fuzzy_matches + forcados)
all_matches = pd.concat([
    exact_match.rename(columns={'TickerBase': 'TickerBase'}),
    fuzzy_df
], ignore_index=True)

# === NOVA RODADA 1: Comparar Ticker (ATIVOS) com TickerBase (COM DINHEIRO) ===
matched_bases = all_matches['TickerBase'].unique()
extra_cd = equity_cd[~equity_cd['TickerBase'].isin(matched_bases)]
extra_at = equity_at[~equity_at['TickerBase'].isin(matched_bases)]

ticker_matches = []
for _, row_cd in extra_cd.iterrows():
    match_row = extra_at[extra_at['Ticker'] == row_cd['TickerBase']]
    if not match_row.empty:
        match_row = match_row.iloc[0]
        ticker_matches.append({
            'Descrição_CD': row_cd['Descrição'],
            'Ticker_CD': row_cd['Ticker'],
            'TickerBase': row_cd['TickerBase'],
            'Classe': row ["Classe"],
            'Quantidade_CD': row_cd['Quantidade'],
            'MarketValue_CD': row_cd['MarketValue'],
            'Nome_MS': match_row['Nome'],
            'Ticker_MS': match_row['Ticker'],
            'Quantidade_MS': match_row['Quantidade'],
            'MarketValue_MS': match_row['MarketValue'],
            'Similaridade': 100
        })

# === NOVA RODADA 2: Comparar primeira palavra de Descrição e Nome (fuzzy) ===
cd_restante = extra_cd[~extra_cd['TickerBase'].isin([m['TickerBase'] for m in ticker_matches])]
at_restante = extra_at[~extra_at['TickerBase'].isin([m['TickerBase'] for m in ticker_matches])]

palavra_matches = []
for _, row_cd in cd_restante.iterrows():
    palavra_cd = str(row_cd['Descrição']).strip().split()[0].upper()
    candidatos = at_restante['Nome'].tolist()
    melhores = process.extract(palavra_cd, candidatos, scorer=fuzz.token_sort_ratio, limit=1)
    if melhores:
        match, score, _ = melhores[0]
        if score >= 80:
            match_row = at_restante[at_restante['Nome'] == match].iloc[0]
            palavra_matches.append({
                'Descrição_CD': row['Descrição'],
                'Ticker_CD': row['Ticker'],
                'TickerBase': row['TickerBase'],
                'Classe': row ["Classe"],
                'Quantidade_CD': row['Quantidade'],
                'MarketValue_CD': row['MarketValue'],
                'Nome_MS': match_row['Nome'],
                'Ticker_MS': match_row['Ticker'],
                'Quantidade_MS': match_row['Quantidade'],
                'MarketValue_MS': match_row['MarketValue'],
                'Similaridade': score
            })

# Junta tudo
extra_df = pd.DataFrame(ticker_matches + palavra_matches)
all_matches = pd.concat([all_matches, extra_df], ignore_index=True)

# Cálculos
if not all_matches.empty:
    all_matches['PrecoUnitario_CD'] = all_matches['MarketValue_CD'] / all_matches['Quantidade_CD']
    all_matches['PrecoUnitario_MS'] = all_matches['MarketValue_MS'] / all_matches['Quantidade_MS']
    all_matches['Diff_Quantidade'] = all_matches['Quantidade_CD'] - all_matches['Quantidade_MS']
    all_matches['Diff_MarketValue'] = all_matches['MarketValue_CD'] - all_matches['MarketValue_MS']
    all_matches['Diff_PrecoUnitario'] = all_matches['PrecoUnitario_CD'] - all_matches['PrecoUnitario_MS']
    all_matches['Pct_Diff_Quantidade'] = all_matches['Diff_Quantidade'] / all_matches['Quantidade_MS']
    all_matches['Pct_Diff_MarketValue'] = all_matches['Diff_MarketValue'] / all_matches['MarketValue_MS']
    all_matches['Pct_Diff_PrecoUnitario'] = all_matches['Diff_PrecoUnitario'] / all_matches['PrecoUnitario_MS']
    all_matches['Destaque'] = np.abs(all_matches['Diff_PrecoUnitario']) > 1

# Marcar os TickerBase usados nos pareamentos
tickersbase_usados_cd = all_matches['TickerBase'].unique()
tickers_usados_at = all_matches['Ticker_MS'].unique()

# Agora removemos da COM DINHEIRO pelo TickerBase
non_matched_cd = equity_cd[~equity_cd['TickerBase'].isin(tickersbase_usados_cd)]

# E removemos da ATIVOS pelo Ticker (não pelo TickerBase!)
non_matched_at = equity_at[~equity_at['Ticker'].isin(tickers_usados_at)]


non_matched_consolidado = pd.concat([
    non_matched_cd.assign(Origem='COM DINHEIRO'),
    non_matched_at.assign(Origem='ATIVOS')
], ignore_index=True)

# Reordenar colunas
if not all_matches.empty:
    col_order = [
        'TickerBase', 'Ticker_MS', 'Ticker_CD', 'Descrição_CD', 'Nome_MS', 
        'Quantidade_CD', 'Quantidade_MS', 'Diff_Quantidade',
        'MarketValue_CD', 'MarketValue_MS', 'Diff_MarketValue',"Classe"]
    all_matches = all_matches[col_order]

# Exportar Excel
wb = Workbook()
wb.remove(wb.active)

ws_pareados = wb.create_sheet("Pareados")
for r in dataframe_to_rows(all_matches, index=False, header=True):
    ws_pareados.append(r)

ws_nao_pareados = wb.create_sheet("Não Pareados")
for r in dataframe_to_rows(non_matched_consolidado, index=False, header=True):
    ws_nao_pareados.append(r)

ws_so_cd = wb.create_sheet("Só COM DINHEIRO")
for r in dataframe_to_rows(non_matched_cd, index=False, header=True):
    ws_so_cd.append(r)

ws_so_at = wb.create_sheet("Só ATIVOS")
for r in dataframe_to_rows(non_matched_at, index=False, header=True):
    ws_so_at.append(r)

if not all_matches.empty:
    diff_market_value_col = col_order.index('Diff_MarketValue') + 1
    for idx, row in enumerate(ws_pareados.iter_rows(min_row=2, max_row=ws_pareados.max_row), 2):
        diff_value = ws_pareados.cell(row=idx, column=diff_market_value_col).value
        if diff_value is not None and abs(diff_value) > 1:
            for cell in row:
                cell.fill = RED_FILL

wb.save("relatorio_consolidado_equity.xlsx")
print("Relatório gerado com sucesso.")
