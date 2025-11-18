import pandas as pd
from src.processor import compute_balance_from_diary_and_ledger


def test_compute_balance_basic():
    diary = pd.DataFrame([
        {'Account': 'Caja', 'Debit': 100, 'Credit': 0},
        {'Account': 'Ventas', 'Debit': 0, 'Credit': 100},
    ])
    ledger = pd.DataFrame([
        {'Account': 'Caja', 'Debit': 50, 'Credit': 0},
        {'Account': 'Bancos', 'Debit': 200, 'Credit': 0},
    ])

    balance = compute_balance_from_diary_and_ledger(diary, ledger)
    # Convert to dict for easy assertions
    bal_map = {r['Account']: r['Balance'] for _, r in balance.iterrows()}
    assert bal_map['Caja'] == 150  # 100 + 50 debits
    assert bal_map['Ventas'] == -100  # credit > debit
    assert bal_map['Bancos'] == 200
