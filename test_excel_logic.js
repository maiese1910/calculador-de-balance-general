// Test script for Excel parsing logic

function normalizeRowKeys(obj) {
    const out = {};
    for (const k of Object.keys(obj)) {
        out[String(k).trim()] = obj[k];
    }
    return out;
}

function parseNumberRaw(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    let s = String(v).trim();
    s = s.replace(/[^0-9,.-]/g, '');
    if (!s) return 0;

    const lastDot = s.lastIndexOf('.');
    const lastComma = s.lastIndexOf(',');

    if (lastDot > -1 && lastComma > -1) {
        if (lastDot > lastComma) {
            // US format: 1,234.56 -> remove commas
            s = s.replace(/,/g, '');
        } else {
            // EU format: 1.234,56 -> remove dots, swap comma to dot
            s = s.replace(/\./g, '').replace(/,/g, '.');
        }
    } else if (lastComma > -1) {
        // Only comma -> treat as decimal separator
        s = s.replace(/,/g, '.');
    }
    // If only dot or no separators, Number() handles it

    const n = Number(s);
    return Number.isNaN(n) ? 0 : n;
}

function getFirstValue(obj, keys) {
    for (const k of keys) {
        if (k in obj && obj[k] !== null && obj[k] !== undefined && String(obj[k]).trim() !== '') return obj[k];
    }
    return null;
}

function computeBalance(diaryRows, ledgerRows) {
    const map = new Map();
    const addRows = rows => {
        rows.forEach(r => {
            const nr = normalizeRowKeys(r);
            // Try common variations for account
            const account = getFirstValue(nr, ['Account', 'Cuenta', 'account', 'cuenta', 'Cuenta nombre', 'Nombre Cuenta', 'NombreCuenta']);
            // also try code + name if account missing
            const codeVal = getFirstValue(nr, ['Código Cta', 'Código Cuenta', 'Codigo Cuenta', 'CodigoCuenta', 'Codigo', 'Account Code', 'AccountCode']);
            const nameVal = getFirstValue(nr, ['Cuenta Contable', 'Nombre Cuenta', 'NombreCuenta', 'Account', 'Cuenta', 'account', 'cuenta']);
            // debit/credit possibilities (include 'Débito Total' / 'Crédito Total')
            const debitRaw = getFirstValue(nr, ['Débito Total', 'Debito Total', 'DebitoTotal', 'Debit', 'Débito', 'Debito', 'Debe', 'DEBE', 'debit', 'debe', 'Amount', 'Importe', 'Monto', 'Valor']);
            const creditRaw = getFirstValue(nr, ['Crédito Total', 'Credito Total', 'CreditoTotal', 'Credit', 'Crédito', 'credito', 'Haber', 'HABER', 'credit', 'haber']);

            let debitVal = parseNumberRaw(debitRaw);
            let creditVal = parseNumberRaw(creditRaw);

            if ((debitRaw !== null) && (creditRaw === null || creditRaw === undefined || String(creditRaw).trim() === '')) {
                const amt = parseNumberRaw(debitRaw);
                if (amt < 0) { debitVal = 0; creditVal = Math.abs(amt); }
                else { debitVal = amt; creditVal = 0; }
            }

            let accountName = account;
            if (!accountName) {
                if (codeVal && nameVal) accountName = `${String(codeVal).trim()} - ${String(nameVal).trim()}`;
                else if (nameVal) accountName = String(nameVal).trim();
                else if (codeVal) accountName = String(codeVal).trim();
            }
            if (!accountName) return;

            // Normalize to uppercase
            const accountKey = accountName.toUpperCase();

            const prev = map.get(accountKey) || { Debit: 0, Credit: 0 };
            prev.Debit += Number(debitVal) || 0;
            prev.Credit += Number(creditVal) || 0;
            map.set(accountKey, prev);
        });
    };

    addRows(diaryRows);
    addRows(ledgerRows);

    const result = [];
    for (const [account, vals] of map.entries()) {
        result.push({ Account: account, Debit: vals.Debit, Credit: vals.Credit, Balance: vals.Debit - vals.Credit });
    }
    result.sort((a, b) => a.Account.localeCompare(b.Account));
    return result;
}

// --- Export Logic Simulation ---
function prepareExportData(balanceData) {
    return balanceData.map(row => ({
        'Cuenta': row.Account,
        'Debe': row.Debit,
        'Haber': row.Credit,
        'Saldo': row.Balance
    }));
}

// Mock data based on user image
const mockDataDiary = [
    {
        "Fecha": "01/11/25",
        "N° Asiento": 1,
        "Código Cta": 1101,
        "Cuenta Contable": "Caja",
        "Debe": "10,000.00",
        "Haber": null
    }
];

const mockDataLedgerFixed = [
    {
        "Cuenta Contable": "1101 - CAJA",
        "Debe": "5,000.00",
        "Haber": null
    }
];

console.log("Running test with mixed casing...");
const result = computeBalance(mockDataDiary, mockDataLedgerFixed);
console.log("Result:", JSON.stringify(result, null, 2));

if (result.length === 1 &&
    result[0].Account === "1101 - CAJA" &&
    result[0].Debit === 15000) {
    console.log("TEST PASSED: Case normalization merged accounts.");
} else {
    console.log("TEST FAILED: Accounts not merged.");
}

console.log("Running test with only Diary...");
const resultDiaryOnly = computeBalance(mockDataDiary, []);
if (resultDiaryOnly.length === 1 && resultDiaryOnly[0].Debit === 10000) {
    console.log("TEST PASSED: Works with only Diary.");
} else {
    console.log("TEST FAILED: Diary only failed.");
}

console.log("Running test for Export Data Preparation...");
const exportData = prepareExportData(result);
console.log("Export Data:", JSON.stringify(exportData, null, 2));

if (exportData.length === 1 &&
    exportData[0].Cuenta === "1101 - CAJA" &&
    exportData[0].Debe === 15000 &&
    exportData[0].Haber === 0 &&
    exportData[0].Saldo === 15000) {
    console.log("TEST PASSED: Export data translated correctly.");
} else {
    console.log("TEST FAILED: Export data translation incorrect.");
}
