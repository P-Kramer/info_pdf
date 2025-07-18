import pandas as pd
import numpy as np
from rapidfuzz import fuzz, process
from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows

# Configurações
RED_FILL = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')

def processar_com_dinheiro(df):
    """Processa a planilha COM DINHEIRO filtrando apenas EQUITYs válidos"""
    # Filtrar linhas relevantes (carteira LA_ e classe EQUITY)
    df = df[
        df['Carteira'].str.startswith('LA_', na=False) & 
        (df['Classe'] == 'EQUITY')
    ].copy()
    
    # Renomear colunas para padrão consistente
    df = df.rename(columns={
        'Ativo': 'Ticker',
        'Saldo Bruto (quant) em 30/06/2025': 'Quantidade',
        'Saldo Bruto em 30/06/2025': 'MarketValue'
    })
    
    # Extrair ticker base (remover prefixos de exchange)
    df['TickerBase'] = df['ticker_cmd_puro'].str.split(':').str[-1].str.strip()
    
    return df[['Descrição', 'Ticker', 'TickerBase', 'Quantidade', 'MarketValue']]

def processar_ativos(df):
    """Processa a planilha Ativos filtrando apenas EQUITYs"""
    # Filtrar apenas equities e renomear colunas
    df = df.rename(columns={
        'Ativo': 'Nome',
        'Quantidade Total': 'Quantidade',
        'Market Value': 'MarketValue'
    })
    
    # Criar ticker base consistente
    df['TickerBase'] = df['Ticker'].str.extract(r'([A-Z]{2,6}$)')[0].fillna(df['Ticker'])
    
    return df[['Nome', 'Ticker', 'TickerBase', 'Quantidade', 'MarketValue']]

# Carregar dados
file_path = 'ativos_extraidos_corrigido_final.xlsx'
df_cd = pd.read_excel(file_path, sheet_name='COM DINHEIRO')
df_at = pd.read_excel(file_path, sheet_name='Ativos')

# Processar dados
equity_cd = processar_com_dinheiro(df_cd)
equity_at = processar_ativos(df_at)

# 1. Matching por TickerBase exato
exact_match = pd.merge(
    equity_cd,
    equity_at,
    on='TickerBase',
    suffixes=('_CD', '_AT'),
    how='inner'
)

# Identificar itens já pareados
matched_tickers = exact_match['TickerBase'].unique()

# 2. Matching fuzzy para itens restantes
remaining_cd = equity_cd[~equity_cd['TickerBase'].isin(matched_tickers)]
remaining_at = equity_at[~equity_at['TickerBase'].isin(matched_tickers)]

fuzzy_matches = []
for _, row in remaining_cd.iterrows():
    # Busca o melhor match na lista de ativos restantes
    match, score, _ = process.extractOne(
        row['Descrição'],
        remaining_at['Nome'].tolist(),
        scorer=fuzz.token_set_ratio
    )
    
    if score > 85:  # Threshold de qualidade
        match_row = remaining_at[remaining_at['Nome'] == match].iloc[0]
        fuzzy_matches.append({
            'Descrição_CD': row['Descrição'],
            'Ticker_CD': row['Ticker'],
            'TickerBase': row['TickerBase'],
            'Quantidade_CD': row['Quantidade'],
            'MarketValue_CD': row['MarketValue'],
            'Nome_AT': match_row['Nome'],
            'Ticker_AT': match_row['Ticker'],
            'Quantidade_AT': match_row['Quantidade'],
            'MarketValue_AT': match_row['MarketValue'],
            'Similaridade': score
        })

# Criar DataFrame de matches fuzzy
fuzzy_df = pd.DataFrame(fuzzy_matches) if fuzzy_matches else pd.DataFrame()

# 3. Juntar todos os matches
all_matches = pd.concat([
    exact_match.rename(columns={'TickerBase': 'TickerBase'}),
    fuzzy_df
], ignore_index=True)

# 4. Calcular métricas de comparação
if not all_matches.empty:
    # Calcular preços unitários
    all_matches['PrecoUnitario_CD'] = all_matches['MarketValue_CD'] / all_matches['Quantidade_CD']
    all_matches['PrecoUnitario_AT'] = all_matches['MarketValue_AT'] / all_matches['Quantidade_AT']
    
    # Calcular diferenças
    all_matches['Diff_Quantidade'] = all_matches['Quantidade_CD'] - all_matches['Quantidade_AT']
    all_matches['Diff_MarketValue'] = all_matches['MarketValue_CD'] - all_matches['MarketValue_AT']
    all_matches['Diff_PrecoUnitario'] = all_matches['PrecoUnitario_CD'] - all_matches['PrecoUnitario_AT']
    
    # Calcular diferenças percentuais
    all_matches['Pct_Diff_Quantidade'] = all_matches['Diff_Quantidade'] / all_matches['Quantidade_AT']
    all_matches['Pct_Diff_MarketValue'] = all_matches['Diff_MarketValue'] / all_matches['MarketValue_AT']
    all_matches['Pct_Diff_PrecoUnitario'] = all_matches['Diff_PrecoUnitario'] / all_matches['PrecoUnitario_AT']
    
    # Flag para diferença > R$1 no preço unitário
    all_matches['Destaque'] = np.abs(all_matches['Diff_PrecoUnitario']) > 1

# 5. Identificar não pareados
non_matched_cd = equity_cd[~equity_cd['TickerBase'].isin(all_matches['TickerBase'])]
non_matched_at = equity_at[~equity_at['TickerBase'].isin(all_matches['TickerBase'])]

# 6. Criar aba consolidada de não pareados
non_matched_consolidado = pd.concat([
    non_matched_cd.assign(Origem='COM DINHEIRO'),
    non_matched_at.assign(Origem='ATIVOS')
], ignore_index=True)

# 7. Reordenar colunas para apresentação
if not all_matches.empty:
    col_order = [
        'TickerBase', 'Ticker_CD', 'Ticker_AT', 
        'Quantidade_CD', 'Quantidade_AT', 'Diff_Quantidade',
        'MarketValue_CD', 'MarketValue_AT', 'Diff_MarketValue']
    all_matches = all_matches[col_order]

# 8. Exportar relatório com formatação
wb = Workbook()
wb.remove(wb.active)  # Remover aba padrão

# Aba de pareados
ws_pareados = wb.create_sheet("Pareados")
for r in dataframe_to_rows(all_matches, index=False, header=True):
    ws_pareados.append(r)

# Aba de não pareados consolidado
ws_nao_pareados = wb.create_sheet("Não Pareados")
for r in dataframe_to_rows(non_matched_consolidado, index=False, header=True):
    ws_nao_pareados.append(r)

# Aba só COM DINHEIRO
ws_so_cd = wb.create_sheet("Só COM DINHEIRO")
for r in dataframe_to_rows(non_matched_cd, index=False, header=True):
    ws_so_cd.append(r)

# Aba só ATIVOS
ws_so_at = wb.create_sheet("Só ATIVOS")
for r in dataframe_to_rows(non_matched_at, index=False, header=True):
    ws_so_at.append(r)

# Aplicar destaque vermelho nas linhas com diferença > R$1 no valor total
if not all_matches.empty:
    # Encontrar posição da coluna Diff_MarketValue
    diff_market_value_col = col_order.index('Diff_MarketValue') + 1
    
    for idx, row in enumerate(ws_pareados.iter_rows(min_row=2, max_row=ws_pareados.max_row), 2):
        diff_value = ws_pareados.cell(row=idx, column=diff_market_value_col).value
        
        # Verificar se é diferente de None e se |valor| > 1
        if diff_value is not None and abs(diff_value) > 1:
            for cell in row:
                cell.fill = RED_FILL

# Salvar relatório
wb.save("relatorio_consolidado_equity.xlsx")
print("Relatório gerado com sucesso: relatorio_consolidado_equity.xlsx")