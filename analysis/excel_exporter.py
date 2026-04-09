import pandas as pd
from utils import fetch_avatar

def export_reports(df: pd.DataFrame, paths: dict, col_order: list):
    """Exporta estritamente para o Excel, assegurando a injeção correta das imagens via xlsxwriter."""
    
    df_temp = df.copy()
    df_temp['FOTO'] = ""
        
    available_columns = [col for col in col_order if col in df_temp.columns]
    df_out = df_temp[available_columns]

    excel_path = paths['excel_out']

    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        df_out.to_excel(writer, sheet_name='Relatório', index=False, header=False, startrow=1)
        
        workbook = writer.book
        worksheet = writer.sheets['Relatório']
        
        format_center = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        format_header = workbook.add_format({
            'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 
            'bold': True, 'bottom': 1, 'bg_color': '#F2F2F2'
        })
        
        worksheet.set_row(0, 40)
        worksheet.freeze_panes(1, 0)
        
        for i, col in enumerate(df_out.columns):
            worksheet.write(0, i, col, format_header)
            
            # Dinâmica de tamanhos de colunas mantida 
            col_width = 14 if col == 'FOTO' else (
                12 if col == '% CONVERSÃO' else (
                    7 if col == 'CPC' else (
                        25 if col == 'OBSERVAÇÕES' else (
                            20 if col == 'CONSULTOR' else max(min(df_out[col].astype(str).str.len().max() + 2, 16), (len(col) // 2) + 5, 11)
                        )
                    )
                )
            )
            worksheet.set_column(i, i, col_width, format_center)
            
        for row_idx, row in enumerate(df_temp.itertuples(), start=1):
            worksheet.set_row(row_idx, 80)
            
            # Lógica otimizada de inserção de imagens do Drive
            if "ui-avatars" not in str(row.FOTO_URL) and row.FOTO_URL:
                try:
                    img_bytes = fetch_avatar(row.FOTO_URL)
                    col_index = df_out.columns.get_loc('FOTO')
                    worksheet.insert_image(
                        row_idx, col_index, 'foto.png', 
                        {'image_data': img_bytes, 'object_position': 1, 'x_offset': 4, 'y_offset': 11}
                    )
                except Exception:
                    pass