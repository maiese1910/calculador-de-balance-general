
// Mocks
const document = {
    getElementById: (id) => {
        if (id === 'validationResults') {
            return {
                innerHTML: '',
                hidden: true,
                className: '',
                appendChild: (child) => {
                    console.log(`[DOM Append] ${child.className}: ${child.innerHTML}`);
                    // validationResults.children.push(child); 
                },
                classList: {
                    add: (cls) => console.log(`[DOM ClassAdd] ${cls}`)
                }
            };
        }
    },
    createElement: (tag) => {
        return { className: '', innerHTML: '' };
    }
};

const formatNumber = (n) => Number(n).toFixed(2);

// LOGIC TO TEST
function validateBooks(diaryRows, ledgerRows) {
    const resultsDiv = document.getElementById('validationResults');
    if (!resultsDiv) return;

    resultsDiv.innerHTML = '';
    resultsDiv.hidden = false;
    resultsDiv.className = 'validation-results';

    let hasErrors = false;
    let hasWarnings = false;

    const addMsg = (msg, type) => {
        const div = document.createElement('div');
        div.className = 'validation-item';
        let icon = 'ℹ️';
        if (type === 'success') icon = '✅';
        if (type === 'warning') icon = '⚠️';
        if (type === 'error') icon = '❌';

        div.innerHTML = `<span>${icon}</span> <span>${msg}</span>`;
        resultsDiv.appendChild(div);

        if (type === 'error') hasErrors = true;
        if (type === 'warning') hasWarnings = true;
    };

    // 1. Validar Libro Diario
    let diaryTotalDebit = 0;
    let diaryTotalCredit = 0;
    if (diaryRows && diaryRows.length > 0) {
        diaryRows.forEach(r => {
            diaryTotalDebit += Number(r.Debit) || 0;
            diaryTotalCredit += Number(r.Credit) || 0;
        });

        const diaryDiff = Math.abs(diaryTotalDebit - diaryTotalCredit);
        if (diaryDiff < 0.01) {
            addMsg(`Libro Diario Cuadrado ($ ${formatNumber(diaryTotalDebit)})`, 'success');
        } else {
            addMsg(`Libro Diario Descuadrado (Dif: $ ${formatNumber(diaryDiff)})`, 'error');
        }
    }

    // 2. Validar Libro Mayor
    let ledgerTotalDebit = 0;
    let ledgerTotalCredit = 0;
    if (ledgerRows && ledgerRows.length > 0) {
        ledgerRows.forEach(r => {
            ledgerTotalDebit += Number(r.Debit) || 0;
            ledgerTotalCredit += Number(r.Credit) || 0;
        });

        const ledgerDiff = Math.abs(ledgerTotalDebit - ledgerTotalCredit);
        if (ledgerDiff < 0.01) {
            addMsg(`Libro Mayor Cuadrado ($ ${formatNumber(ledgerTotalDebit)})`, 'success');
        } else {
            addMsg(`Libro Mayor Descuadrado (Dif: $ ${formatNumber(ledgerDiff)})`, 'error');
        }
    }

    // 3. Validar Consistencia (Si ambos existen)
    if (diaryRows && diaryRows.length > 0 && ledgerRows && ledgerRows.length > 0) {
        const diffDebit = Math.abs(diaryTotalDebit - ledgerTotalDebit);
        const diffCredit = Math.abs(diaryTotalCredit - ledgerTotalCredit);

        if (diffDebit < 1 && diffCredit < 1) { // Tolerancia $1 por redondeos
            addMsg('Consistencia Correcta: Totales de Diario y Mayor coinciden.', 'success');
        } else {
            addMsg(`Inconsistencia: Los totales del Diario difieren del Mayor.`, 'warning');
            if (diffDebit >= 1) addMsg(`- Diferencia en Débitos: $ ${formatNumber(diffDebit)}`, 'warning');
            if (diffCredit >= 1) addMsg(`- Diferencia en Créditos: $ ${formatNumber(diffCredit)}`, 'warning');
        }
    }

    if (hasErrors) resultsDiv.classList.add('error');
    else if (hasWarnings) resultsDiv.classList.add('warning');
    else resultsDiv.classList.add('success');
}

// TEST CASES

console.log('--- TEST 1: Both Balanced and Consistent ---');
validateBooks(
    [{ Debit: 100, Credit: 0 }, { Debit: 0, Credit: 100 }],
    [{ Debit: 100, Credit: 0 }, { Debit: 0, Credit: 100 }]
);

console.log('\n--- TEST 2: Diary Unbalanced ---');
validateBooks(
    [{ Debit: 100, Credit: 0 }, { Debit: 0, Credit: 90 }], // Unbalanced
    [{ Debit: 100, Credit: 0 }, { Debit: 0, Credit: 100 }]
);

console.log('\n--- TEST 3: Inconsistent (Diario != Mayor) ---');
validateBooks(
    [{ Debit: 100, Credit: 0 }, { Debit: 0, Credit: 100 }],
    [{ Debit: 200, Credit: 0 }, { Debit: 0, Credit: 200 }] // Different totals
);
