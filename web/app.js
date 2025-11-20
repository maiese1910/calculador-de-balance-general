// app.js - Lógica del frontend para leer Excel, calcular balances y (opcional) guardar en Firebase
(() => {
  // Corrección para Electron: intentar cargar xlsx vía require primero
  if (typeof require !== 'undefined') {
    try {
      const localXLSX = require('xlsx');
      if (localXLSX) {
        window.XLSX = localXLSX;
        console.log('XLSX cargado vía require (Electron)');
      }
    } catch (e) {
      console.warn('No se pudo cargar xlsx vía require:', e);
    }
  }

  // Verificar si XLSX está disponible
  if (typeof XLSX === 'undefined') {
    console.error('La librería XLSX no está cargada.');
  }

  // ========================================
  // SISTEMA DE LICENCIAS
  // ========================================

  // Verificar licencia al cargar la página
  function checkLicenseOnStartup() {
    if (!hasValidLicense()) {
      showLicenseModal();
    } else {
      hideLicenseModal();
    }
  }

  function showLicenseModal() {
    const modal = document.getElementById('licenseModal');
    if (modal) {
      modal.hidden = false;
      modal.style.display = 'flex';
    }
  }

  function hideLicenseModal() {
    const modal = document.getElementById('licenseModal');
    if (modal) {
      modal.hidden = true;
      modal.style.display = 'none';
    }
  }

  // Configurar eventos del modal de licencia
  function setupLicenseModal() {
    const input = document.getElementById('licenseInput');
    const btn = document.getElementById('activateBtn');
    const error = document.getElementById('licenseError');

    if (!input || !btn || !error) {
      console.error('Elementos del modal de licencia no encontrados');
      return;
    }

    // Formatear input automáticamente
    input.addEventListener('input', (e) => {
      let value = e.target.value.toUpperCase().replace(/[^A-Z0-9]/g, '');
      let formatted = '';
      for (let i = 0; i < value.length && i < 16; i++) {
        if (i > 0 && i % 4 === 0) formatted += '-';
        formatted += value[i];
      }
      e.target.value = formatted;
    });

    // Manejar activación
    btn.addEventListener('click', () => {
      const key = input.value.trim().toUpperCase();

      if (validateLicenseKey(key)) {
        saveLicense(key);
        hideLicenseModal();
        error.hidden = true;
        alert('✅ Licencia activada correctamente');
      } else {
        error.textContent = '❌ Código de licencia inválido. Verifica el formato.';
        error.hidden = false;
      }
    });

    // Permitir activar con Enter
    input.addEventListener('keypress', (e) => {
      if (e.key === 'Enter') {
        btn.click();
      }
    });
  }

  // Inicializar sistema de licencias cuando el DOM esté listo
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', () => {
      checkLicenseOnStartup();
      setupLicenseModal();
    });
  } else {
    checkLicenseOnStartup();
    setupLicenseModal();
  }

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

  // Función para leer un archivo Excel y convertirlo a JSON
  function readExcelFile(file) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onload = e => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = XLSX.read(data, { type: 'array' });
          // Depuración: mostrar nombres de las hojas
          console.log('Hojas del libro:', workbook.SheetNames);
          setStatus(`Hojas detectadas: ${workbook.SheetNames.join(', ')}`, 'info');

          const firstName = workbook.SheetNames[0];
          const ws = workbook.Sheets[firstName];
          // Intentar analizar a JSON usando la fila de encabezado
          let json = XLSX.utils.sheet_to_json(ws, { defval: null });

          // Fallback: si no devuelve filas, intentar analizar como matriz de matrices y construir objetos
          if ((!json || json.length === 0)) {
            const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
            // buscar la primera fila que parezca un encabezado (tiene algunos valores no nulos)
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
                // construir objeto
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
          console.error('Error parseando Excel:', err);
          reject(err);
        }
      };
      fr.onerror = (err) => {
        console.error('Error de lectura de archivo (FileReader):', err);
        reject(err);
      };
      fr.readAsArrayBuffer(file);
    });
  }

  // Normalizar claves de fila (eliminar espacios en nombres de columnas)
  function normalizeRowKeys(obj) {
    const out = {};
    for (const k of Object.keys(obj)) {
      out[String(k).trim()] = obj[k];
    }
    return out;
  }

  // Formateador de números usando la configuración regional del usuario (muestra 2 decimales)
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

  // ========================================
  // VALIDACIÓN DE BALANCE CUADRADO
  // ========================================

  /**
   * Verifica si el balance está cuadrado (Total Debe = Total Haber)
   * @param {Array} balance - Array de objetos con Debit y Credit
   * @returns {Object} - {balanced: boolean, totalDebit, totalCredit, difference, message}
   */
  function checkBalanced(balance) {
    let totalDebit = 0;
    let totalCredit = 0;

    balance.forEach(row => {
      totalDebit += row.Debit || 0;
      totalCredit += row.Credit || 0;
    });

    const difference = Math.abs(totalDebit - totalCredit);
    const tolerance = 0.01; // Tolerancia de 1 centavo

    if (difference < tolerance) {
      return {
        balanced: true,
        totalDebit,
        totalCredit,
        difference: 0,
        message: '✅ Balance cuadrado'
      };
    } else {
      return {
        balanced: false,
        totalDebit,
        totalCredit,
        difference,
        message: `⚠️ Balance descuadrado - Diferencia: ${formatNumber(difference)}`
      };
    }
  }

  /**
   * Muestra el estado del balance (cuadrado/descuadrado)
   */
  function showBalanceStatus(balanceCheck) {
    const statusDiv = document.getElementById('balanceStatus');
    if (!statusDiv) return;

    statusDiv.hidden = false;
    statusDiv.className = 'balance-status';

    if (balanceCheck.balanced) {
      statusDiv.classList.add('balanced');
      statusDiv.innerHTML = `
        <div>${balanceCheck.message}</div>
        <div class="balance-status-details">
          Total Debe: ${formatNumber(balanceCheck.totalDebit)} | 
          Total Haber: ${formatNumber(balanceCheck.totalCredit)}
        </div>
      `;
    } else {
      statusDiv.classList.add('unbalanced');
      statusDiv.innerHTML = `
        <div>${balanceCheck.message}</div>
        <div class="balance-status-details">
          Total Debe: ${formatNumber(balanceCheck.totalDebit)} | 
          Total Haber: ${formatNumber(balanceCheck.totalCredit)}
        </div>
      `;
    }
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

  // Analizador de números robusto: maneja cadenas como "1.234,56", "1,234.56", "$ 1.234,56", etc.
  function parseNumberRaw(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number') return v;
    let s = String(v).trim();
    // eliminar símbolos de moneda y espacios
    s = s.replace(/[^0-9,.-]/g, '');
    if (!s) return 0;

    const lastDot = s.lastIndexOf('.');
    const lastComma = s.lastIndexOf(',');

    if (lastDot > -1 && lastComma > -1) {
      if (lastDot > lastComma) {
        // Formato US: 1,234.56 -> eliminar comas
        s = s.replace(/,/g, '');
      } else {
        // Formato EU: 1.234,56 -> eliminar puntos, cambiar coma a punto
        s = s.replace(/\./g, '').replace(/,/g, '.');
      }
    } else if (lastComma > -1) {
      // Solo coma -> tratar como separador decimal (común en entradas simples en ES)
      s = s.replace(/,/g, '.');
    }
    // Si solo hay punto o no hay separadores, Number() lo maneja (formato US por defecto)

    const n = Number(s);
    return Number.isNaN(n) ? 0 : n;
  }

  function getFirstValue(obj, keys) {
    for (const k of keys) {
      if (k in obj && obj[k] !== null && obj[k] !== undefined && String(obj[k]).trim() !== '') return obj[k];
    }
    return null;
  }

  // Función principal para calcular el balance
  function computeBalance(diaryRows, ledgerRows) {
    const map = new Map();
    const addRows = rows => {
      rows.forEach(r => {
        const nr = normalizeRowKeys(r);
        // Intentar variaciones comunes para la cuenta
        const account = getFirstValue(nr, ['Account', 'Cuenta', 'account', 'cuenta', 'Cuenta nombre', 'Nombre Cuenta', 'NombreCuenta']);
        // también intentar código + nombre si falta la cuenta
        const codeVal = getFirstValue(nr, ['Código Cta', 'Código Cuenta', 'Codigo Cuenta', 'CodigoCuenta', 'Codigo', 'Account Code', 'AccountCode']);
        const nameVal = getFirstValue(nr, ['Cuenta Contable', 'Nombre Cuenta', 'NombreCuenta', 'Account', 'Cuenta', 'account', 'cuenta']);
        // posibilidades de débito/crédito (incluir 'Débito Total' / 'Crédito Total')
        const debitRaw = getFirstValue(nr, ['Débito Total', 'Debito Total', 'DebitoTotal', 'Debit', 'Débito', 'Debito', 'Debe', 'DEBE', 'debit', 'debe', 'Amount', 'Importe', 'Monto', 'Valor']);
        const creditRaw = getFirstValue(nr, ['Crédito Total', 'Credito Total', 'CreditoTotal', 'Credit', 'Crédito', 'credito', 'Haber', 'HABER', 'credit', 'haber']);

        let debitVal = parseNumberRaw(debitRaw);
        let creditVal = parseNumberRaw(creditRaw);

        // Si hay una sola columna de monto con valores firmados y el crédito es cero, dividir por signo
        if ((debitRaw !== null) && (creditRaw === null || creditRaw === undefined || String(creditRaw).trim() === '')) {
          // tratar debitRaw como monto firmado
          const amt = parseNumberRaw(debitRaw);
          if (amt < 0) { debitVal = 0; creditVal = Math.abs(amt); }
          else { debitVal = amt; creditVal = 0; }
        }

        // construir cadena de cuenta si falta pero existen código/nombre
        let accountName = account;
        if (!accountName) {
          if (codeVal && nameVal) accountName = `${String(codeVal).trim()} - ${String(nameVal).trim()}`;
          else if (nameVal) accountName = String(nameVal).trim();
          else if (codeVal) accountName = String(codeVal).trim();
        }
        if (!accountName) return;
        // Normalizar a mayúsculas para fusionar "Caja" y "CAJA"
        const accountKey = accountName.toUpperCase();

        const prev = map.get(accountKey) || { Debit: 0, Credit: 0 };
        prev.Debit += Number(debitVal) || 0;
        prev.Credit += Number(creditVal) || 0;
        // Almacenar la clave en mayúsculas para consistencia
        map.set(accountKey, prev);
      });
    };

    addRows(diaryRows);
    addRows(ledgerRows);

    const result = [];
    for (const [account, vals] of map.entries()) {
      result.push({ Account: account, Debit: vals.Debit, Credit: vals.Credit, Balance: vals.Debit - vals.Credit });
    }
    // Ordenar por Cuenta
    result.sort((a, b) => a.Account.localeCompare(b.Account));
    return result;
  }

  function renderResult(rows) {
    resultTbody.innerHTML = '';
    // limpiar pie de página
    if (resultTfoot) resultTfoot.innerHTML = '';
    rows.forEach(r => {
      const tr = document.createElement('tr');
      const accountTd = document.createElement('td'); accountTd.textContent = r.Account; tr.appendChild(accountTd);
      const debitTd = document.createElement('td'); debitTd.textContent = formatNumber(r.Debit); tr.appendChild(debitTd);
      const creditTd = document.createElement('td'); creditTd.textContent = formatNumber(r.Credit); tr.appendChild(creditTd);
      const balTd = document.createElement('td'); balTd.textContent = formatNumber(r.Balance); tr.appendChild(balTd);
      // colorear por signo
      if (Number(r.Balance) > 0) balTd.classList.add('balance-positive');
      else if (Number(r.Balance) < 0) balTd.classList.add('balance-negative');
      else balTd.classList.add('balance-zero');
      resultTbody.appendChild(tr);
    });
    noResult.hidden = true;
    resultTable.hidden = false;

    // Mostrar totales si está habilitado
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
    if (!diaryFile && !ledgerFile) {
      alert('Seleccione al menos un archivo: libro diario o libro mayor');
      return;
    }
    try {
      computeBtn.disabled = true;
      setStatus('Leyendo archivos...', 'info');
      const diaryRows = diaryFile ? await readExcelFile(diaryFile) : [];
      const ledgerRows = ledgerFile ? await readExcelFile(ledgerFile) : [];
      console.log('diaryRows', diaryRows.slice(0, 5));
      console.log('ledgerRows', ledgerRows.slice(0, 5));
      // mostrar conteos
      setStatus(`Leído libro diario: ${diaryRows.length} filas. Libro mayor: ${ledgerRows.length} filas.`, 'info');
      if (!diaryRows.length && !ledgerRows.length) {
        setStatus('No se encontraron filas legibles en los archivos seleccionados.', 'error');
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

      // Validar si el balance está cuadrado
      const balanceCheck = checkBalanced(balance);
      showBalanceStatus(balanceCheck);

      // habilitar guardar si firebase está configurado
      saveBtn.disabled = !(window.firebaseConfig && window.firebaseConfig.projectId);
    } catch (err) {
      console.error(err);
      setStatus('Error leyendo los archivos: ' + (err.message || err), 'error');
      alert('Error leyendo los archivos: ' + err.message);
    } finally {
      computeBtn.disabled = false;
    }
  });

  exportBtn.addEventListener('click', async () => {
    if (!lastBalance) return;

    // 1. Traducir datos y preparar para exportación
    const dataForExport = lastBalance.map(row => ({
      'Cuenta': row.Account,
      'Debe': row.Debit,
      'Haber': row.Credit,
      'Saldo': row.Balance
    }));

    // 2. Crear hoja de trabajo
    const ws = XLSX.utils.json_to_sheet(dataForExport);

    // 3. Ajustar anchos de columna (visual "más ordenado")
    const wscols = [
      { wch: 50 }, // Cuenta (ancho generoso para nombres largos)
      { wch: 15 }, // Debe
      { wch: 15 }, // Haber
      { wch: 15 }  // Saldo
    ];
    ws['!cols'] = wscols;

    // 4. Aplicar formato de números a las celdas de montos
    const range = XLSX.utils.decode_range(ws['!ref']);
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
      for (let C = 1; C <= 3; ++C) {
        const cell_address = { c: C, r: R };
        const cell_ref = XLSX.utils.encode_cell(cell_address);
        if (!ws[cell_ref]) continue;

        ws[cell_ref].z = '#,##0.00';
        ws[cell_ref].t = 'n';
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Balance');

    // Detectar si estamos en Electron
    const isElectron = typeof require !== 'undefined' && typeof process !== 'undefined' && process.versions && process.versions.electron;

    if (isElectron) {
      // En Electron: usar diálogo de guardado
      try {
        const { dialog } = require('electron').remote || require('@electron/remote');
        const path = require('path');

        const result = await dialog.showSaveDialog({
          title: 'Guardar Balance General',
          defaultPath: path.join(require('os').homedir(), 'Downloads', 'balance_general.xlsx'),
          filters: [
            { name: 'Excel Files', extensions: ['xlsx'] },
            { name: 'All Files', extensions: ['*'] }
          ]
        });

        if (!result.canceled && result.filePath) {
          XLSX.writeFile(wb, result.filePath);
          alert('Archivo guardado exitosamente en: ' + result.filePath);
        }
      } catch (err) {
        console.error('Error usando diálogo de Electron:', err);
        // Fallback: guardar en Descargas
        const path = require('path');
        const os = require('os');
        const filePath = path.join(os.homedir(), 'Downloads', 'balance_general.xlsx');
        XLSX.writeFile(wb, filePath);
        alert('Archivo guardado en: ' + filePath);
      }
    } else {
      // En navegador: descarga normal
      XLSX.writeFile(wb, 'balance_general.xlsx');
    }
  });

  // Inicialización de Firebase (compat)
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

  // Intentar inicializar Firebase al cargar
  initFirebaseIfConfigured();

})();
