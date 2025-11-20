// app.js - lógica del frontend para leer Excel, calcular balances y (opcional) guardar en Firebase
(() => {
  const diaryInput = document.getElementById('diaryFile');
  const ledgerInput = document.getElementById('ledgerFile');
  const computeBtn = document.getElementById('computeBtn');
  const exportBtn = document.getElementById('exportBtn');
  const saveBtn = document.getElementById('saveBtn');
  const resultTable = document.getElementById('resultTable');
  const resultTbody = resultTable.querySelector('tbody');
  const resultTfoot = resultTable.querySelector('tfoot');
  const statusArea = document.getElementById('statusArea');
  const noResult = document.getElementById('noResult');
  const currencyToggle = document.getElementById('currencyToggle');
  const currencySymbol = document.getElementById('currencySymbol');
  const totalsToggle = document.getElementById('totalsToggle');

  let lastBalance = null;

  function readExcelFile(file) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = e => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          // Debug: expose sheet names
          console.log('Workbook sheets:', workbook.SheetNames);
          setStatus(`Hojas detectadas: ${workbook.SheetNames.join(', ')}`, 'info');

          const firstName = workbook.SheetNames[0];
          const ws = workbook.Sheets[firstName];
          // Try parsing to JSON using header row
          let json = XLSX.utils.sheet_to_json(ws, { defval: null });

          // Fallback: if no rows returned, try parse as array-of-arrays and build objects
          if ((!json || json.length === 0)) {
            const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
            // find first row that looks like a header (has some non-null values)
            let headerRowIndex = -1;
            for (let i = 0; i < Math.min(5, aoa.length); i++) {
              const row = aoa[i];
              if (row && row.some(v => v !== null && String(v).trim() !== '')) {
                headerRowIndex = i;
                break;
              }
            }
            if (headerRowIndex >= 0) {
              const headers = aoa[headerRowIndex].map(h => h === null ? '' : String(h).trim());
              const rows = [];
              for (let r = headerRowIndex + 1; r < aoa.length; r++) {
                const row = aoa[r];
                if (!row) continue;
                // build object
                const obj = {};
                let any = false;
                for (let c = 0; c < headers.length; c++) {
                  const key = headers[c] || `col${c}`;
                  const val = row[c] !== undefined ? row[c] : null;
                  obj[key] = val;
                  if (val !== null && String(val).trim() !== '') any = true;
                }
                if (any) rows.push(obj);
              }
              json = rows;
            }
          }

          resolve(json);
        } catch (err) {
          reject(err);
        }
      };
      fr.onerror = reject;
      fr.readAsArrayBuffer(file);
    });
  }

  function normalizeRowKeys(obj) {
    const out = {};
    for (const k of Object.keys(obj)) {
      out[String(k).trim()] = obj[k];
    }
    return out;
  }

  // Number formatter using user's locale (shows 2 decimals)
  const nf = new Intl.NumberFormat(navigator.language || 'es-ES', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  function formatNumber(val) {
    if (val === null || val === undefined) return '';
    const n = Number(val);
    if (Number.isNaN(n)) return '';
    const formatted = nf.format(n);
    if (currencyToggle && currencyToggle.checked) {
      const sym = (currencySymbol && currencySymbol.value) ? currencySymbol.value : '$';
      return sym + ' ' + formatted;
    }
    return formatted;
  }

  function setStatus(msg, type = 'info') {
    if (!statusArea) return;
    statusArea.textContent = msg;
    statusArea.classList.remove('info', 'error', 'success');
    statusArea.classList.add(type);
  }

  const diaryPreview = document.getElementById('diaryPreview');
  const ledgerPreview = document.getElementById('ledgerPreview');
  function showPreview(elem, rows) {
    if (!elem) return;
    if (!rows || !rows.length) {
      elem.hidden = true;
      return;
    }
    const small = rows.slice(0, 5).map(r => normalizeRowKeys(r));
    elem.textContent = JSON.stringify(small, null, 2);
    elem.hidden = false;
  }

  // Robust number parser: handle strings like "1.234,56", "1,234.56", "$ 1.234,56", etc.
  function parseNumberRaw(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    let s = String(v).trim();
    // remove currency symbols and spaces
    s = s.replace(/[^0-9,.-]/g, '');
    if (!s) return 0;
    // If contains both '.' and ',', assume '.' thousands and ',' decimal (common in ES locales)
    if (s.indexOf('.') > -1 && s.indexOf(',') > -1) {
      // remove dots, replace comma with dot
      s = s.replace(/\./g, '').replace(/,/g, '.');
    } else if (s.indexOf(',') > -1 && s.indexOf('.') === -1) {
      // only comma -> treat as decimal separator
      s = s.replace(/,/g, '.');
    } else {
      // only dots or only digits -> leave as is (dots are decimal or thousands, Number will parse)
    }
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
        const account = getFirstValue(nr, ['Account','Cuenta','account','cuenta','Cuenta contable','CuentaContable','Cuenta nombre','Nombre Cuenta','NombreCuenta']);
        // also try code + name if account missing
        const codeVal = getFirstValue(nr, ['Código Cuenta','Codigo Cuenta','CodigoCuenta','Codigo','Account Code','AccountCode']);
        const nameVal = getFirstValue(nr, ['Nombre Cuenta','NombreCuenta','Account','Cuenta','account','cuenta']);
        // debit/credit possibilities (include 'Débito Total' / 'Crédito Total')
        const debitRaw = getFirstValue(nr, ['Débito Total','Debito Total','DebitoTotal','Debit','Débito','Debito','Debe','DEBE','debit','debe','Amount','Importe','Monto','Valor']);
        const creditRaw = getFirstValue(nr, ['Crédito Total','Credito Total','CreditoTotal','Credit','Crédito','credito','Haber','HABER','credit','haber']);

        let debitVal = parseNumberRaw(debitRaw);
        let creditVal = parseNumberRaw(creditRaw);

        // If there is a single amount column with signed values and credit is zero, split by sign
        if ((debitRaw !== null) && (creditRaw === null || creditRaw === undefined || String(creditRaw).trim() === '')) {
          // treat debitRaw as signed amount
          const amt = parseNumberRaw(debitRaw);
          if (amt < 0) { debitVal = 0; creditVal = Math.abs(amt); }
          else { debitVal = amt; creditVal = 0; }
        }

        // build account string if missing but code/name exist
        let accountName = account;
        if (!accountName) {
          if (codeVal && nameVal) accountName = `${String(codeVal).trim()} - ${String(nameVal).trim()}`;
          else if (nameVal) accountName = String(nameVal).trim();
          else if (codeVal) accountName = String(codeVal).trim();
        }
        if (!accountName) return;
        const prev = map.get(account) || { Debit: 0, Credit: 0 };
        prev.Debit += Number(debitVal) || 0;
        prev.Credit += Number(creditVal) || 0;
        map.set(accountName, prev);
      });
    };

    addRows(diaryRows);
    addRows(ledgerRows);

    const result = [];
    for (const [account, vals] of map.entries()) {
      result.push({ Account: account, Debit: vals.Debit, Credit: vals.Credit, Balance: vals.Debit - vals.Credit });
    }
    // Sort by Account
    result.sort((a, b) => a.Account.localeCompare(b.Account));
    return result;
  }

  function renderResult(rows) {
    resultTbody.innerHTML = '';
    // clear footer
    if (resultTfoot) resultTfoot.innerHTML = '';
    rows.forEach(r => {
      const tr = document.createElement('tr');
      const accountTd = document.createElement('td'); accountTd.textContent = r.Account; tr.appendChild(accountTd);
      const debitTd = document.createElement('td'); debitTd.textContent = formatNumber(r.Debit); tr.appendChild(debitTd);
      const creditTd = document.createElement('td'); creditTd.textContent = formatNumber(r.Credit); tr.appendChild(creditTd);
      const balTd = document.createElement('td'); balTd.textContent = formatNumber(r.Balance); tr.appendChild(balTd);
      // color by sign
      if (Number(r.Balance) > 0) balTd.classList.add('balance-positive');
      else if (Number(r.Balance) < 0) balTd.classList.add('balance-negative');
      else balTd.classList.add('balance-zero');
      resultTbody.appendChild(tr);
    });
    noResult.hidden = true;
    resultTable.hidden = false;

    // Show totals if enabled
    if (totalsToggle && totalsToggle.checked && rows.length > 0) {
      const totals = rows.reduce((acc, r) => {
        acc.Debit += Number(r.Debit) || 0;
        acc.Credit += Number(r.Credit) || 0;
        acc.Balance += Number(r.Balance) || 0;
        return acc;
      }, { Debit: 0, Credit: 0, Balance: 0 });

      if (resultTfoot) {
        const tr = document.createElement('tr');
        tr.classList.add('totals-row');
        const th = document.createElement('th'); th.textContent = 'Totales'; th.colSpan = 1; tr.appendChild(th);
        const debitTd = document.createElement('td'); debitTd.textContent = formatNumber(totals.Debit); tr.appendChild(debitTd);
        const creditTd = document.createElement('td'); creditTd.textContent = formatNumber(totals.Credit); tr.appendChild(creditTd);
        const balTd = document.createElement('td'); balTd.textContent = formatNumber(totals.Balance); tr.appendChild(balTd);
        if (Number(totals.Balance) > 0) balTd.classList.add('balance-positive');
        else if (Number(totals.Balance) < 0) balTd.classList.add('balance-negative');
        else balTd.classList.add('balance-zero');
        resultTfoot.appendChild(tr);
      }
    }
  }

  computeBtn.addEventListener('click', async () => {
    const diaryFile = diaryInput.files[0];
    const ledgerFile = ledgerInput.files[0];
    if (!diaryFile || !ledgerFile) {
      alert('Seleccione ambos archivos: libro diario y libro mayor');
      return;
    }
    try {
      computeBtn.disabled = true;
      setStatus('Leyendo archivos...', 'info');
      const diaryRows = await readExcelFile(diaryFile);
      const ledgerRows = await readExcelFile(ledgerFile);
      console.log('diaryRows', diaryRows.slice(0,5));
      console.log('ledgerRows', ledgerRows.slice(0,5));
      // show counts
      setStatus(`Leído libro diario: ${diaryRows.length} filas. Libro mayor: ${ledgerRows.length} filas.`, 'info');
      if (!diaryRows.length || !ledgerRows.length) {
        setStatus('Uno de los archivos no contiene filas legibles. Verifique el formato y que la hoja contiene datos.', 'error');
        exportBtn.disabled = true;
        saveBtn.disabled = true;
        return;
      }

      setStatus('Calculando balance...', 'info');
      const balance = computeBalance(diaryRows, ledgerRows);
      lastBalance = balance;
      if (!balance || !balance.length) {
        setStatus('No se encontraron cuentas tras el cálculo. Compruebe los encabezados de columnas (Account / Cuenta / Cuenta contable).', 'error');
        exportBtn.disabled = true;
        saveBtn.disabled = true;
        return;
      }
      renderResult(balance);
      exportBtn.disabled = false;
      setStatus(`Balance generado: ${balance.length} cuentas.`, 'success');
      // enable save if firebase configured
      saveBtn.disabled = !(window.firebaseConfig && window.firebaseConfig.projectId);
    } catch (err) {
      console.error(err);
      setStatus('Error leyendo los archivos: ' + (err.message || err), 'error');
      alert('Error leyendo los archivos: ' + err.message);
    } finally {
      computeBtn.disabled = false;
    }
  });

  exportBtn.addEventListener('click', () => {
    if (!lastBalance) return;
    const ws = XLSX.utils.json_to_sheet(lastBalance);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Balance');
    XLSX.writeFile(wb, 'balance_general.xlsx');
  });

  // Firebase init (compat)
  function initFirebaseIfConfigured() {
    try {
      if (window.firebaseConfig && window.firebaseConfig.projectId) {
        firebase.initializeApp(window.firebaseConfig);
        window._firestore = firebase.firestore();
        saveBtn.disabled = false;
      }
    } catch (err) {
      console.warn('Firebase no inicializado:', err);
    }
  }

  saveBtn.addEventListener('click', async () => {
    if (!lastBalance) return alert('No hay balance para guardar');
    if (!window._firestore) return alert('Firebase no está configurado. Consulte README.');
    try {
      saveBtn.disabled = true;
      const doc = {
        createdAt: new Date().toISOString(),
        balance: lastBalance
      };
      const ref = await window._firestore.collection('balances').add(doc);
      alert('Balance guardado con id: ' + ref.id);
    } catch (err) {
      console.error(err);
      alert('Error guardando en Firestore: ' + err.message);
    } finally {
      saveBtn.disabled = false;
    }
  });

  // Try to initialize Firebase on load
  initFirebaseIfConfigured();

})();
