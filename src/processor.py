import pandas as pd

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza los nombres de las columnas a los estándares 'Account', 'Debit', 'Credit'."""
    # Mapa de nombres estándar a posibles variaciones (en minúsculas)
    column_mapping = {
        'Account': [
            'account', 'cuenta', 'nombre cuenta', 'nombrecuenta', 'código cuenta', 'codigocuenta',
            'cuenta contable', 'cuentacontable', 'cuenta nombre',
        ],
        'Debit': [
            'debit', 'debe', 'débito', 'debito', 'debitos', 'débitos',
            'débito total', 'debito total', 'debitototal', 'amount', 'importe', 'monto', 'valor'
        ],
        'Credit': [
            'credit', 'haber', 'crédito', 'credito', 'creditos', 'créditos',
            'crédito total', 'credito total', 'creditototal'
        ],
    }

    # Crear una versión más limpia de las columnas: eliminar espacios y convertir a minúsculas
    clean_cols = {c: str(c).strip().lower() for c in df.columns}
    
    # Verificar el patrón de combinación Código + Nombre
    code_variations = ['código cuenta', 'codigo cuenta', 'codigocuenta', 'codigo', 'código cta', 'account code']
    name_variations = ['cuenta contable', 'nombre cuenta', 'nombrecuenta', 'account name']
    
    code_col = next((orig for orig, clean in clean_cols.items() if clean in code_variations), None)
    name_col = next((orig for orig, clean in clean_cols.items() if clean in name_variations), None)
    
    df_processed = df.copy()
    
    # Si ambos están presentes, combinarlos y eliminar los originales para evitar colisiones
    if code_col and name_col:
        df_processed['Account'] = df_processed[code_col].astype(str).str.strip() + ' - ' + df_processed[name_col].astype(str).str.strip()
        df_processed = df_processed.drop(columns=[code_col, name_col])
        # Recalcular columnas limpias para las columnas restantes
        clean_cols = {c: str(c).strip().lower() for c in df_processed.columns}
    
    # Invertir el mapeo para búsqueda
    lookup = {}
    for std, variations in column_mapping.items():
        for v in variations:
            lookup[v] = std
            
    # Renombrar columnas
    new_names = {}
    for original, clean in clean_cols.items():
        if clean in lookup:
            new_names[original] = lookup[clean]
            
    return df_processed.rename(columns=new_names)

def compute_balance_from_diary_and_ledger(diary_df: pd.DataFrame, ledger_df: pd.DataFrame) -> pd.DataFrame:
    """Calcula un balance general simple a partir de DataFrames de diario y mayor.

    Expectativas:
    - Ambos DataFrames contienen una columna `Account` (o alias) y columnas numéricas `Debit`/`Credit` (o alias).
    - La función devuelve un DataFrame con `Account` y `Balance` (Debit - Credit) agregados.
    """
    # Normalizar nombres de columnas
    df_d = normalize_columns(diary_df)
    df_l = normalize_columns(ledger_df)

    # Asegurar que existan las columnas requeridas
    for df in (df_d, df_l):
        if 'Account' not in df.columns:
            # ¿Intentar encontrar una columna que parezca una cuenta si no se encuentra por nombre?
            # Por ahora, estricto en tener al menos una columna mapeada para Cuenta
            raise ValueError(f'El DataFrame debe contener una columna "Cuenta" (o similar). Columnas encontradas: {list(df.columns)}')
        
        # Rellenar Debit/Credit faltantes con ceros
        for col in ('Debit', 'Credit'):
            if col not in df.columns:
                df[col] = 0
            else:
                # Asegurar numérico
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Concatenar y agrupar por Cuenta
    combined = pd.concat([df_d[['Account', 'Debit', 'Credit']], df_l[['Account', 'Debit', 'Credit']]], ignore_index=True)
    agg = combined.groupby('Account', dropna=False).sum(numeric_only=True)
    agg['Balance'] = agg['Debit'] - agg['Credit']
    result = agg.reset_index()[['Account', 'Debit', 'Credit', 'Balance']]
    
    # Renombrar a español para la salida
    result = result.rename(columns={
        'Account': 'Cuenta',
        'Debit': 'Debe',
        'Credit': 'Haber',
        'Balance': 'Saldo'
    })
    return result

def export_balance_to_excel(balance_df: pd.DataFrame, path: str) -> None:
    """Exporta el DataFrame de balance a un archivo Excel."""
    # Usar ExcelWriter para aplicar formatos
    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        balance_df.to_excel(writer, index=False, sheet_name='Balance')
        
        # Obtener la hoja de trabajo y el libro
        workbook = writer.book
        worksheet = writer.sheets['Balance']
        
        # Formatos
        # Nota: openpyxl no usa un objeto 'Format' como xlsxwriter, se modifica la celda directamente o se usa NamedStyle
        # Pero para simplicidad en pandas, podemos iterar o ajustar columnas.
        
        # Ajustar anchos de columna
        worksheet.column_dimensions['A'].width = 50  # Cuenta
        worksheet.column_dimensions['B'].width = 15  # Debe
        worksheet.column_dimensions['C'].width = 15  # Haber
        worksheet.column_dimensions['D'].width = 15  # Saldo
        
        # Aplicar formato de número a las columnas B, C, D
        # Iterar sobre las filas de datos (asumiendo que empiezan en la fila 2)
        from openpyxl.styles import NamedStyle
        
        number_style = NamedStyle(name='number_style', number_format='#,##0.00')
        
        # Iterar sobre las columnas de datos (B, C, D son índices 2, 3, 4 en openpyxl 1-based)
        for col_idx in range(2, 5): # 2, 3, 4
            col_letter = chr(64 + col_idx) # B, C, D
            for cell in worksheet[col_letter]:
                if cell.row > 1: # Saltar encabezado
                    cell.number_format = '#,##0.00'

