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
    bal_map = {r['Cuenta']: r['Saldo'] for _, r in balance.iterrows()}
    assert bal_map['Caja'] == 150  # 100 + 50 debits
    assert bal_map['Ventas'] == -100  # credit > debit
    assert bal_map['Bancos'] == 200

def test_compute_balance_spanish_columns():
    diary = pd.DataFrame([
        {'Cuenta': 'Caja', 'Debe': 100, 'Haber': 0},
    ])
    ledger = pd.DataFrame([
        {'Nombre Cuenta': 'Caja', 'Débito': 50, 'Crédito': 0},
    ])
    
    balance = compute_balance_from_diary_and_ledger(diary, ledger)
    bal_map = {r['Cuenta']: r['Saldo'] for _, r in balance.iterrows()}
    assert bal_map['Caja'] == 150

def test_compute_balance_messy_columns():
    diary = pd.DataFrame([
        {'  cuEnta ': 'Caja', ' DEBE': 100}, # Missing Credit
    ])
    ledger = pd.DataFrame([
        {'account': 'Caja', ' credit ': 20}, # Missing Debit
    ])
    
    balance = compute_balance_from_diary_and_ledger(diary, ledger)
    bal_map = {r['Cuenta']: r['Saldo'] for _, r in balance.iterrows()}
    assert bal_map['Caja'] == 80 # 100 - 20

def test_compute_balance_missing_account_error():
    diary = pd.DataFrame([{'Foo': 1}])
    ledger = pd.DataFrame([{'Bar': 1}])
    try:
        compute_balance_from_diary_and_ledger(diary, ledger)
        assert False, "Should have raised ValueError"
    except ValueError as e:
        assert "debe contener una columna \"Cuenta\"" in str(e)

def test_compute_balance_cuenta_contable():
    diary = pd.DataFrame([
        {'Cuenta Contable': 'Caja', 'Debe': 100, 'Haber': 0},
    ])
    ledger = pd.DataFrame([
        {'CuentaContable': 'Caja', 'Débito': 50, 'Crédito': 0},
    ])
    
    balance = compute_balance_from_diary_and_ledger(diary, ledger)
    bal_map = {r['Cuenta']: r['Saldo'] for _, r in balance.iterrows()}
    assert bal_map['Caja'] == 150
