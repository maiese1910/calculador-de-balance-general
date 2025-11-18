// app.js - lógica del frontend para leer Excel, calcular balances y (opcional) guardar en Firebase
(() => {
  const diaryInput = document.getElementById('diaryFile');
  const ledgerInput = document.getElementById('ledgerFile');
  const computeBtn = document.getElementById('computeBtn');
  const exportBtn = document.getElementById('exportBtn');
  const saveBtn = document.getElementById('saveBtn');
  const resultTable = document.getElementById('resultTable');
  const resultTbody = resultTable.querySelector('tbody');
  const noResult = document.getElementById('noResult');

  let lastBalance = null;

  function readExcelFile(file) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = e => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          const firstName = workbook.SheetNames[0];
          const ws = workbook.Sheets[firstName];
          const json = XLSX.utils.sheet_to_json(ws, { defval: null });
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

  function computeBalance(diaryRows, ledgerRows) {
    const map = new Map();
    const addRows = rows => {
      rows.forEach(r => {
        const nr = normalizeRowKeys(r);
        // Try common variations
        const account = nr['Account'] ?? nr['Cuenta'] ?? nr['account'] ?? nr['cuenta'];
        const debit = Number(nr['Debit'] ?? nr['Débito'] ?? nr['Debito'] ?? nr['debit'] ?? 0) || 0;
        const credit = Number(nr['Credit'] ?? nr['Crédito'] ?? nr['credito'] ?? nr['credit'] ?? 0) || 0;
        if (!account) return;
        const prev = map.get(account) || { Debit: 0, Credit: 0 };
        prev.Debit += debit;
        prev.Credit += credit;
        map.set(account, prev);
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
    rows.forEach(r => {
      const tr = document.createElement('tr');
      const accountTd = document.createElement('td'); accountTd.textContent = r.Account; tr.appendChild(accountTd);
      const debitTd = document.createElement('td'); debitTd.textContent = r.Debit; tr.appendChild(debitTd);
      const creditTd = document.createElement('td'); creditTd.textContent = r.Credit; tr.appendChild(creditTd);
      const balTd = document.createElement('td'); balTd.textContent = r.Balance; tr.appendChild(balTd);
      resultTbody.appendChild(tr);
    });
    noResult.hidden = true;
    resultTable.hidden = false;
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
      const diaryRows = await readExcelFile(diaryFile);
      const ledgerRows = await readExcelFile(ledgerFile);
      const balance = computeBalance(diaryRows, ledgerRows);
      lastBalance = balance;
      renderResult(balance);
      exportBtn.disabled = false;
      // enable save if firebase configured
      saveBtn.disabled = !(window.firebaseConfig && window.firebaseConfig.projectId);
    } catch (err) {
      console.error(err);
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
