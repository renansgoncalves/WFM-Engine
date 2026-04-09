import os
from config import PATHS, COL_ORDER
from core import WFMProcessor
from excel_exporter import export_reports
from bi_exporter import export_to_bi
from utils import format_to_ms

def main():
    required_keys = ['engagements', 'breaks', 'consultores_info']
    if not all(os.path.exists(PATHS[key]) for key in required_keys):
        print("[!] Erro: Arquivos CSV de entrada não encontrados.")
        return

    print(">> Calculando métricas de WFM...")
    wfm = WFMProcessor()
    
    # 1. Carregamento e limpeza de bases
    df_info = wfm.process_info(PATHS['consultores_info'])
    df_eng = wfm.clean_engagements(PATHS['engagements'])
    df_brk = wfm.clean_breaks(PATHS['breaks'], df_info)

    # 2. Processamento profundo de métricas
    df = wfm.get_metrics(df_eng, df_brk, df_info)

    df['% ENGANO'] = (df['ENGANO'] / df['ACIONAMENTOS PRODUTIVOS'] * 100).fillna(0).astype(int).astype(str) + "%"
    df['% CONVERSÃO'] = (df['STATUS POSITIVOS'] / df['PROPOSTAS'] * 100).fillna(0).astype(int).astype(str) + "%"
    df['EQUIPE'] = df['EQUIPE_INFO']

    # 3. Exportação para Power BI
    # Exportamos a base aqui para que o Power BI acesse os decimais brutos
    export_to_bi(df, PATHS)

    # 4. Formatações visuais (convertendo os raw decimals para strings formatadas)
    time_columns = ['TEMPO EM LIGAÇÃO', 'TEMPO NÃO TABELADO', 'TEMPO DE OCIOSIDADE', 'ALMOÇO', 'BANHEIRO', 'TEMPO TOTAL DE PAUSA']
    for col in time_columns: 
        raw_col_name = f"{col if 'TOTAL' not in col else 'TEMPO TOTAL DE PAUSA'}_raw"
        df[col] = df[raw_col_name].apply(format_to_ms)
    
    def construct_avatar_url(drive_id):
        if drive_id:
            return f"https://drive.google.com/uc?export=view&id={drive_id}"
        return "https://ui-avatars.com/api/?name=User&format=png"
        
    df['FOTO_URL'] = df['DRIVE_ID'].apply(construct_avatar_url)
    df['OBSERVAÇÕES'] = ""

    # 5. Exportação para relatório Excel
    export_reports(df, PATHS, COL_ORDER)
    print(f"OK: Relatório Excel gerado em: {os.path.dirname(PATHS['excel_out'])}")

if __name__ == "__main__":
    main()