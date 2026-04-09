import pandas as pd
import os

def export_to_bi(df: pd.DataFrame, paths: dict):
    """
    Gera um arquivo plano (CSV) projetado para o Power BI.
    Mantém os números flutuantes e colunas limpas para evitar conversões de string erradas no DAX.
    """
    df_bi = df.copy()
    
    # Tratamento da média de tempo de ligações CPC para consumo otimizado no Power BI
    if 'MEDIA_TEMPO_CPC_raw' in df_bi.columns:
        df_bi['MEDIA TEMPO CPC (MINUTOS)'] = df_bi['MEDIA_TEMPO_CPC_raw'].round(2)
    
    # Exportação padronizada
    bi_target = paths['bi_out']
    df_bi.to_csv(bi_target, index=False, encoding='utf-8-sig', sep=';', decimal=',')
    print(f"OK: Base de dados Power BI exportada em: {bi_target}")