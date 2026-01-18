
// Mocks needed for standardizeRows
function getFirstValue(row, keys) {
    for (const k of keys) {
        if (row[k] !== undefined && row[k] !== null && row[k] !== '') return row[k];
    }
    return null;
}

function normalizeRowKeys(row) {
    const out = {};
    Object.keys(row).forEach(k => {
        const cleanKey = String(k).trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        out[cleanKey] = row[k];
        // Keep original too just in case
        out[String(k).trim()] = row[k];
    });
    return out;
}

function standardizeRows(rows) {
    if (!rows || rows.length === 0) return [];

    const firstObj = normalizeRowKeys(rows[0]);
    const keys = Object.keys(firstObj);
    console.log('Keys detectadas en fila 0:', keys);

    return rows.map(r => {
        const nr = normalizeRowKeys(r);
        // Debug
        // console.log('Procesando fila:', JSON.stringify(r));

        const out = {
            Date: getFirstValue(nr, ['Fecha', 'Date', 'Day', 'Dia']),
            Code: getFirstValue(nr, ['Código', 'Codigo', 'Code', 'Cuenta Numero', 'Numero Cuenta']),
            Account: getFirstValue(nr, ['Cuenta', 'Account', 'Nombre Cuenta', 'Descripcion', 'Detalle']),
            Debit: getFirstValue(nr, ['Debe', 'Debit', 'Débito', 'Debito']),
            Credit: getFirstValue(nr, ['Haber', 'Credit', 'Crédito', 'Credito'])
        };

        const vals = Object.values(nr);
        // Heuristic fallback logic (copied from app.js)
        if (!out.Account && !out.Debit && !out.Credit && vals.length >= 3) {
            if (vals.length >= 5) {
                out.Date = vals[0];
                out.Code = vals[1];
                out.Account = vals[2];
                out.Debit = vals[3];
                out.Credit = vals[4];
            } else if (vals.length === 4) {
                out.Code = vals[0];
                out.Account = vals[1];
                out.Debit = vals[2];
                out.Credit = vals[3];
            } else if (vals.length === 3) {
                out.Account = vals[0];
                out.Debit = vals[1];
                out.Credit = vals[2];
            }
        }
        return out;
    });
}

// MIMIC THE USER'S EXCEL READ
// Based on the screenshot:
// Row 4 (index 3 presumably if 0-indexed) has headers.
// Col A: Código, Col B: Empty, Col C: Cuenta, Col D: Debe, Col E: Haber, Col F: Saldo
// When XLSX reads this with sheet_to_json, it might produce objects like:
// { "Código": 101, "__EMPTY": null, "Cuenta": "BANCOS", "Debe": 14000, "Haber": 2500, "Saldo": 11500 }
// OR if using header:1 (AoA):
// [ "Código", null, "Cuenta", "Debe", "Haber", "Saldo" ]

// Scenario A: JSON objects from XLSX (assuming it successfully found headers)
const mockJsonFromXlsx = [
    { "Código": 101, "Cuenta": "BANCOS", "Debe": 14000, "Haber": 2500, "Saldo": 11500 }, // __EMPTY might be skipped by xlsx default
    { "Código": 105, "Cuenta": "EQUIPO DE COMPUTO", "Debe": 2000, "Haber": 0, "Saldo": 2000 }
];

console.log('--- TEST SCENARIO A (Clean JSON) ---');
const processedA = standardizeRows(mockJsonFromXlsx);
console.log(processedA);


// Scenario B: User's layout might cause correct headers NOT to be found if they are not in row 0
// readExcelFile logic attempts to find a header row.
// Let's simulate the AoA (Array of Arrays) that happens before sheet_to_json if the first parse fails or as raw data
const rawAoA = [
    ["LIBRO MAYOR"],
    [null],
    ["MOVIMIENTOS POR CUENTA"],
    ["Código", null, "Cuenta", "Debe", "Haber", "Saldo"], // Header Row (Index 3)
    [101, null, "BANCOS", 14000, 2500, 11500],
    [105, null, "EQUIPO DE COMPUTO", 2000, 0, 2000]
];

// Logic from readExcelFile to find header
function simulateReadExcelFileLogic(aoa) {
    let headerRowIndex = -1;
    for (let i = 0; i < Math.min(5, aoa.length); i++) {
        const row = aoa[i];
        // Heuristic: row has some string content?
        // Logic in app.js: row && row.some(v => v !== null && String(v).trim() !== '')
        // But specifically we want the one with "Código" or "Cuenta" probably.
        // The current app.js just takes the FIRST row that has meaningful content.
        // ERROR POTENTIAL: "LIBRO MAYOR" is in row 0. It might pick row 0 as header.

        // Let's see what the current code does:
        if (row && row.some(v => v !== null && String(v).trim() !== '')) {
            // It picks THE FIRST ONE.
            // If Row 0 is "LIBRO MAYOR", it picks ["LIBRO MAYOR"] as headers.
            // Then Row 1 is null.
            // Row 3 is "MOVIMIENTOS..."
            // Row 4 is proper headers.

            // If it picks Row 0, then the keys are ['LIBRO MAYOR'].
            // Subsequent rows:
            // [101, null, "BANCOS"...] -> { "LIBRO MAYOR": 101, "__EMPTY": null, "__EMPTY_1": "BANCOS"... }
            // This fails standardizeRows lookup.
            headerRowIndex = i;
            break;
        }
    }

    console.log('Header Row Found at Index:', headerRowIndex);
    if (headerRowIndex >= 0) {
        const headers = aoa[headerRowIndex].map(h => h === null ? '' : String(h).trim());
        console.log('Headers:', headers);

        const rows = [];
        for (let r = headerRowIndex + 1; r < aoa.length; r++) {
            const row = aoa[r];
            if (!row) continue;
            const obj = {};
            let any = false;
            for (let c = 0; c < row.length; c++) { // headers might be shorter if "LIBRO MAYOR" is len 1
                const key = headers[c] || `col${c}`;
                const val = row[c] !== undefined ? row[c] : null;
                obj[key] = val;
                if (val !== null && String(val).trim() !== '') any = true;
            }
            if (any) rows.push(obj);
        }
        return rows;
    }
    return [];
}

console.log('\n--- TEST SCENARIO B (Raw AoA Parsing Logic) ---');
const parsedB = simulateReadExcelFileLogic(rawAoA);
console.log('Parsed Rows (Before Standardize):', JSON.stringify(parsedB.slice(0, 3), null, 2));

const processedB = standardizeRows(parsedB);
console.log('Standardized B:', processedB);
// Expectation: This will fail to map 'Account' or 'Code' because keys will be wrong (e.g. 'LIBRO MAYOR' or 'colX')
