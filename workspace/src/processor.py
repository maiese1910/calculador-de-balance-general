import pandas as pd

def compute_balance_from_diary_and_ledger(diary_df: pd.DataFrame, ledger_df: pd.DataFrame) -> pd.DataFrame:
    """Compute a simple balance general from diary and ledger DataFrames.

    Expectations (initial convention):
    - Both DataFrames contain an `Account` column and `Debit` and `Credit` numeric columns.
    - Function returns a DataFrame with `Account` and `Balance` (Debit - Credit) aggregated.

    This is a starting point: classification into assets/liabilities/equity can be added later.
    """
    # Normalize column names
    df_d = diary_df.rename(columns={c: c.strip() for c in diary_df.columns})
    df_l = ledger_df.rename(columns={c: c.strip() for c in ledger_df.columns})

    # Ensure required columns exist
    for df in (df_d, df_l):
        if 'Account' not in df.columns:
            raise ValueError('DataFrame must contain an "Account" column')
        # Fill missing Debit/Credit with zeros
        for col in ('Debit', 'Credit'):
            if col not in df.columns:
                df[col] = 0

    # Concatenate and group by Account
    combined = pd.concat([df_d[['Account', 'Debit', 'Credit']], df_l[['Account', 'Debit', 'Credit']]], ignore_index=True)
    agg = combined.groupby('Account', dropna=False).sum(numeric_only=True)
    agg['Balance'] = agg['Debit'] - agg['Credit']
    result = agg.reset_index()[['Account', 'Debit', 'Credit', 'Balance']]
    return result

def export_balance_to_excel(balance_df: pd.DataFrame, path: str) -> None:
    """Export the balance DataFrame to an Excel file."""
    balance_df.to_excel(path, index=False)
