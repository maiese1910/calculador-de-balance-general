
// Mock dependencies
const normalizeRowKeys = (obj) => {
    const out = {};
    for (const k of Object.keys(obj)) out[String(k).trim()] = obj[k];
    return out;
};
const getFirstValue = (obj, keys) => {
    for (const k of keys) if (k in obj && obj[k] !== null && String(obj[k]).trim() !== '') return obj[k];
    return null;
};
const parseNumberRaw = (v) => Number(v) || 0;

// Paste computeBalance here (from Step 152) to test it
function computeBalance(diaryRows, ledgerRows) {
    const map = new Map();

    const addRows = rows => {
        rows.forEach(r => {
            let account = r.Account;
            const code = r.Code;
            if (!account && !code) return;
            let key = account;
            if (code) {
                const c = String(code).trim();
                if (!account) key = c + ' - (Sin Nombre)';
                else if (!String(account).includes(c)) key = c + ' - ' + String(account).trim();
            }
            const finalKey = String(key || '').toUpperCase();
            const prev = map.get(finalKey) || { Debit: 0, Credit: 0 };
            prev.Debit += Number(r.Debit) || 0;
            prev.Credit += Number(r.Credit) || 0;
            map.set(finalKey, prev);
        });
    };

    addRows(diaryRows);
    addRows(ledgerRows);

    const classifyAccount = (name, codeStr) => {
        const n = name.trim().toLowerCase();
        const match = n.match(/^(\d+)/);
        let c = codeStr || (match ? match[1] : '');

        if (c) {
            const firstDigit = parseInt(c[0]);
            if (firstDigit >= 1 && firstDigit <= 3) return 'REAL';
            if (firstDigit >= 4) return 'NOMINAL';
        }

        if (n.includes('inventario') || n.includes('mercaderia') || n.includes('existencia')) {
            if (n.includes('costo') && !n.includes('por cobrar')) return 'NOMINAL';
            return 'REAL';
        }

        if (
            n.includes('venta') || n.includes('ingreso') || n.includes('gasto') ||
            n.includes('compra') || n.includes('costo') || n.includes('sueldo') ||
            n.includes('salario') || n.includes('servicio') || n.includes('honorario') ||
            n.includes('depreciacion') || n.includes('amortizacion') ||
            n.includes('intereses pagados') || n.includes('flete')
        ) {
            if (n.includes('por pagar') || n.includes('pagado por anticipado')) return 'REAL';
            return 'NOMINAL';
        }
        return 'REAL';
    };

    const trialBalance = [];
    const nominalAccounts = [];

    const balanceSheet = {
        assets: { current: [], nonCurrent: [] },
        liabilities: { current: [], nonCurrent: [] },
        equity: []
    };

    let incomeStatementDebits = 0;
    let incomeStatementCredits = 0;

    for (const [accountKey, vals] of map.entries()) {
        const bal = vals.Debit - vals.Credit;
        const separatorIdx = accountKey.indexOf(' - ');
        let code = '';
        let name = accountKey;
        if (separatorIdx > -1) {
            code = accountKey.substring(0, separatorIdx);
            name = accountKey.substring(separatorIdx + 3);
        }

        const row = { Account: accountKey, Debit: vals.Debit, Credit: vals.Credit, Balance: bal, Code: code, CleanName: name };
        trialBalance.push(row);

        const type = classifyAccount(name, code);

        if (type === 'NOMINAL') {
            nominalAccounts.push(row);
            incomeStatementDebits += vals.Debit;
            incomeStatementCredits += vals.Credit;
        } else {
            const n = name.toLowerCase();
            let group = 'ASSET';

            if (code && (code.startsWith('2'))) group = 'LIABILITY';
            else if (code && (code.startsWith('3'))) group = 'EQUITY';
            else {
                if (n.includes('capital') || n.includes('utilidad') || n.includes('superavit') || n.includes('reserva') || n.includes('patrimonio') || n.includes('resultado acumulado')) group = 'EQUITY';
                else if (n.includes('pasivo') || n.includes('por pagar') || n.includes('proveedor') || n.includes('obligacion') || n.includes('anticipo de clientes')) group = 'LIABILITY';
                else group = 'ASSET';
            }

            if (group === 'ASSET' && (n.includes('banco') || n.includes('cuenta corriente'))) {
                if (row.Balance < 0) {
                    group = 'LIABILITY';
                    row.Account = row.Account + ' (Sobregiro)';
                    row.Balance = Math.abs(row.Balance);
                    const temp = row.Debit; row.Debit = row.Credit; row.Credit = temp;
                }
            }

            if (group === 'ASSET') {
                if (
                    n.includes('caja') || n.includes('banco') || n.includes('cliente') || n.includes('por cobrar') ||
                    n.includes('inventario') || n.includes('mercaderia') || n.includes('iva') || n.includes('anticipo') ||
                    (code && code.startsWith('11'))
                ) {
                    balanceSheet.assets.current.push(row);
                } else {
                    balanceSheet.assets.nonCurrent.push(row);
                }
            } else if (group === 'LIABILITY') {
                if (n.includes('largo plazo') || n.includes('hipoteca') || (code && code.startsWith('22'))) {
                    balanceSheet.liabilities.nonCurrent.push(row);
                } else {
                    balanceSheet.liabilities.current.push(row);
                }
            } else {
                balanceSheet.equity.push(row);
            }
        }
    }

    const netIncome = incomeStatementCredits - incomeStatementDebits;

    const netIncomeRow = {
        Account: 'UTILIDAD (PÉRDIDA) DEL EJERCICIO',
        Debit: netIncome < 0 ? Math.abs(netIncome) : 0,
        Credit: netIncome > 0 ? netIncome : 0,
        Balance: -netIncome
    };

    balanceSheet.equity.push(netIncomeRow);

    const financialPositionFlat = [
        ...balanceSheet.assets.current,
        ...balanceSheet.assets.nonCurrent,
        ...balanceSheet.liabilities.current,
        ...balanceSheet.liabilities.nonCurrent,
        ...balanceSheet.equity
    ];

    return { trialBalance, incomeStatement: nominalAccounts, financialPosition: financialPositionFlat, netIncome };
}

// TEST CASES
const rows = [
    { Account: 'Caja', Code: '1105', Debit: 1000, Credit: 0 },
    { Account: 'Banco', Code: '1110', Debit: 0, Credit: 500 }, // Sobregiro! Balance -500
    { Account: 'Ventas', Code: '4135', Debit: 0, Credit: 2000 },
    { Account: 'Gastos Diversos', Code: '5200', Debit: 800, Credit: 0 },
    { Account: 'Capital', Code: '3100', Debit: 0, Credit: 1000 } // Total Haber: 3500. Total Debe: 1800. Unbalanced but testing logic.
    // Net Income = 2000 - 800 = 1200.
    // Assets: Caja 1000.
    // Liabilities: Banco 500.
    // Equity: Capital 1000 + Income 1200 = 2200.
    // A = 1000. P+E = 2700. (Original inputs unbalanced).
];

console.log('Running logic...');
try {
    const res = computeBalance(rows, []);
    console.log('Success!', res);
    console.log('NetIncome:', res.netIncome);
    console.log('BankRow:', res.financialPosition.find(r => r.CleanName.includes('Banco')));
} catch (e) {
    console.error('Crash!', e);
    process.exit(1);
}
