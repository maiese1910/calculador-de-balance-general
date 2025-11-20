
import pandas as pd
import sys
import os

# Add src to path
sys.path.append(os.path.join(os.getcwd(), 'src'))

from processor import normalize_columns

def test_normalize_columns_combined():
    data = {
        'Código Cta': ['1101', '3101'],
        'Cuenta Contable': ['Caja', 'Capital Social'],
        'Debe': [10000, 0],
        'Haber': [0, 10000]
    }
    df = pd.DataFrame(data)
    
    print("Original DataFrame:")
    print(df)
    
    normalized = normalize_columns(df)
    
    print("\nNormalized DataFrame:")
    print(normalized)
    
    assert 'Account' in normalized.columns
    assert normalized.iloc[0]['Account'] == '1101 - Caja'
    assert normalized.iloc[1]['Account'] == '3101 - Capital Social'
    assert 'Código Cta' not in normalized.columns
    assert 'Cuenta Contable' not in normalized.columns
    
    print("\nTEST PASSED: Combined columns correctly.")

def test_normalize_columns_legacy():
    data = {
        'Cuenta': ['Caja', 'Capital'],
        'Debe': [100, 0],
        'Haber': [0, 100]
    }
    df = pd.DataFrame(data)
    normalized = normalize_columns(df)
    assert 'Account' in normalized.columns
    assert normalized.iloc[0]['Account'] == 'Caja'
    print("\nTEST PASSED: Legacy columns correctly handled.")

if __name__ == "__main__":
    try:
        test_normalize_columns_combined()
        test_normalize_columns_legacy()
    except Exception as e:
        print(f"\nTEST FAILED: {e}")
        sys.exit(1)
