import pandas as pd

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names to standard 'Account', 'Debit', 'Credit'."""
    # Map of standard names to possible variations (lowercase)
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

    # Create a cleaner version of columns: strip whitespace and lowercase
    clean_cols = {c: str(c).strip().lower() for c in df.columns}
    
    # Invert mapping for lookup
    lookup = {}
    for std, variations in column_mapping.items():
        for v in variations:
            lookup[v] = std
            
    # Rename columns
    new_names = {}
    for original, clean in clean_cols.items():
        if clean in lookup:
            new_names[original] = lookup[clean]
            
    return df.rename(columns=new_names)

def compute_balance_from_diary_and_ledger(diary_df: pd.DataFrame, ledger_df: pd.DataFrame) -> pd.DataFrame:
    """Compute a simple balance general from diary and ledger DataFrames.

    Expectations:
    - Both DataFrames contain an `Account` column (or alias) and `Debit`/`Credit` numeric columns (or aliases).
    - Function returns a DataFrame with `Account` and `Balance` (Debit - Credit) aggregated.
    """
    # Normalize column names
    df_d = normalize_columns(diary_df)
    df_l = normalize_columns(ledger_df)

    # Ensure required columns exist
    for df in (df_d, df_l):
        if 'Account' not in df.columns:
            # Try to find a column that looks like an account if not found by name? 
            # For now, strict on having at least one mapped column for Account
            raise ValueError(f'El DataFrame debe contener una columna "Cuenta" (o similar). Columnas encontradas: {list(df.columns)}')
        
        # Fill missing Debit/Credit with zeros
        for col in ('Debit', 'Credit'):
            if col not in df.columns:
                df[col] = 0
            else:
                # Ensure numeric
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

    # Concatenate and group by Account
    combined = pd.concat([df_d[['Account', 'Debit', 'Credit']], df_l[['Account', 'Debit', 'Credit']]], ignore_index=True)
    agg = combined.groupby('Account', dropna=False).sum(numeric_only=True)
    agg['Balance'] = agg['Debit'] - agg['Credit']
    result = agg.reset_index()[['Account', 'Debit', 'Credit', 'Balance']]
    
    # Rename to Spanish for output
    result = result.rename(columns={
        'Account': 'Cuenta',
        'Debit': 'Debe',
        'Credit': 'Haber',
        'Balance': 'Saldo'
    })
    return result

def export_balance_to_excel(balance_df: pd.DataFrame, path: str) -> None:
    """Export the balance DataFrame to an Excel file."""
    balance_df.to_excel(path, index=False)
