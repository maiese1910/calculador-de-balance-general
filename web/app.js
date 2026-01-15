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

  function showSuccessNotification() {
    // Crear overlay
    const overlay = document.createElement('div');
    overlay.className = 'success-overlay';

    // Crear notificación
    const notification = document.createElement('div');
    notification.className = 'success-notification';
    notification.innerHTML = `
      <div class="icon">🎉</div>
      <h3>¡Licencia Activada!</h3>
      <p>Tu aplicación está lista para usar</p>
    `;

    document.body.appendChild(overlay);
    document.body.appendChild(notification);

    // Remover después de 2.5 segundos
    setTimeout(() => {
      notification.style.animation = 'fadeOut 0.3s ease-out forwards';
      overlay.style.animation = 'fadeOut 0.3s ease-out forwards';
      setTimeout(() => {
        document.body.removeChild(notification);
        document.body.removeChild(overlay);
      }, 300);
    }, 2500);
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
        showSuccessNotification();
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

    // Clasificación de cuentas
    // 1, 2, 3 -> Reales (Situación Financiera)
    // 4, 5, 6, etc -> Nominales (Estado de Resultados)
    // Si no hay código, intentar por nombre
    const classifyAccount = (name) => {
      const n = name.trim();
      // Intentar detectar código numérico al inicio
      const match = n.match(/^(\d+)/);
      if (match) {
        const firstDigit = parseInt(match[1][0]);
        if (firstDigit >= 1 && firstDigit <= 3) return 'REAL'; // Activo, Pasivo, Patrimonio
        if (firstDigit >= 4) return 'NOMINAL'; // Ingresos, Gastos, Costos
      }
      
      // Fallback por palabras clave
      const lower = n.toLowerCase();
      if (lower.includes('capital') || lower.includes('acumulada') || lower.includes('banco') || lower.includes('caja') || lower.includes('activo') || lower.includes('pasivo') || lower.includes('patrimonio') || lower.includes('por pagar') || lower.includes('por cobrar')) return 'REAL';
      if (lower.includes('venta') || lower.includes('ingreso') || lower.includes('gasto') || lower.includes('costo') || lower.includes('sueldo') || lower.includes('servicio') || lower.includes('honorario') || lower.includes('depreciacion')) return 'NOMINAL';
      
      return 'REAL'; // Default conservador
    };

    const trialBalance = [];
    let nominalAccounts = [];
    let realAccounts = [];

    // Totales para Estado de Resultados
    let incomeStatementDebits = 0;
    let incomeStatementCredits = 0;

    for (const [account, vals] of map.entries()) {
      const bal = vals.Debit - vals.Credit;
      const row = { Account: account, Debit: vals.Debit, Credit: vals.Credit, Balance: bal };
      
      // 1. Balance de Comprobación (Todos)
      trialBalance.push(row);

      // 2. Clasificar
      const type = classifyAccount(account);
      if (type === 'NOMINAL') {
        nominalAccounts.push(row);
        incomeStatementDebits += vals.Debit;
        incomeStatementCredits += vals.Credit;
      } else {
        realAccounts.push(row);
      }
    }

    // Ordenar
    trialBalance.sort((a, b) => a.Account.localeCompare(b.Account));
    nominalAccounts.sort((a, b) => a.Account.localeCompare(b.Account));
    realAccounts.sort((a, b) => a.Account.localeCompare(b.Account));

    // Calcular Utilidad/Pérdida del Ejercicio
    // Ingresos (Crédito) - Egresos (Débito)
    // Si Créditos > Débitos = Utilidad (Saldo Acreedor)
    const netIncome = incomeStatementCredits - incomeStatementDebits;
    
    // Crear objeto para la Utilidad en el Estado de Situación Financiera
    // Se agrega al Patrimonio (Haber si es utilidad)
    const netIncomeRow = {
      Account: 'UTILIDAD (PÉRDIDA) DEL EJERCICIO',
      Debit: netIncome < 0 ? Math.abs(netIncome) : 0, 
      Credit: 0,
      Balance: -netIncome 
    };
    
    if (netIncome >= 0) {
        netIncomeRow.Credit = netIncome;
    } else {
        netIncomeRow.Debit = Math.abs(netIncome);
    }

    // Agregar Utilidad a Cuentas Reales (Situación Financiera)
    const financialPosition = [...realAccounts, netIncomeRow];
    financialPosition.sort((a, b) => a.Account.localeCompare(b.Account));

    return {
      trialBalance,
      incomeStatement: nominalAccounts,
      financialPosition,
      netIncome
    };
  }

  // Helper para renderizar una tabla específica
  function createTableHTML(title, rows, showTotals = true) {
    if (!rows || !rows.length) return `<div class="table-section"><h3>${title}</h3><p>No hay cuentas.</p></div>`;
    
    const totals = rows.reduce((acc, r) => {
        acc.Debit += Number(r.Debit) || 0;
        acc.Credit += Number(r.Credit) || 0;
        acc.Balance += Number(r.Balance) || 0;
        return acc;
    }, { Debit: 0, Credit: 0, Balance: 0 });

    let html = `
      <div class="table-section">
        <h3>${title}</h3>
        <table class="result-table">
          <thead>
            <tr>
              <th>Cuenta</th>
              <th>Debe</th>
              <th>Haber</th>
              <th>Saldo</th>
            </tr>
          </thead>
          <tbody>
    `;

    rows.forEach(r => {
      let balClass = '';
      if (Number(r.Balance) > 0) balClass = 'balance-positive';
      else if (Number(r.Balance) < 0) balClass = 'balance-negative';
      else balClass = 'balance-zero';

      html += `
        <tr>
          <td>${r.Account}</td>
          <td>${formatNumber(r.Debit)}</td>
          <td>${formatNumber(r.Credit)}</td>
          <td class="${balClass}">${formatNumber(r.Balance)}</td>
        </tr>
      `;
    });

    html += `</tbody>`;

    if (showTotals) {
      let balTotalClass = '';
      if (totals.Balance > 0) balTotalClass = 'balance-positive';
      else if (totals.Balance < 0) balTotalClass = 'balance-negative';
      else balTotalClass = 'balance-zero';

      html += `
          <tfoot>
            <tr class="totals-row">
              <th>Totales</th>
              <td>${formatNumber(totals.Debit)}</td>
              <td>${formatNumber(totals.Credit)}</td>
              <td class="${balTotalClass}">${formatNumber(totals.Balance)}</td>
            </tr>
          </tfoot>
      `;
    }
    
    html += `</table></div>`;
    return html;
  }

  function renderResult(data) {
    const container = document.getElementById('resultTablesContainer');
    if (!container) {
        // Si no existe el contenedor específico, usar resultArea o limpiarlo
        // Vamos a crear dinámicamente o asumir que reemplazamos el resultTable original
        // Por compatibilidad, ocultamos la tabla original y usaremos un div nuevo
    }
    
    // Ocultar elementos viejos
    if (resultTbody) resultTbody.innerHTML = '';
    if (resultTfoot) resultTfoot.innerHTML = '';
    resultTable.hidden = true;
    noResult.hidden = true;

    // Buscar o crear contenedor de reportes
    let reportContainer = document.getElementById('reportContainer');
    if (!reportContainer) {
        reportContainer = document.createElement('div');
        reportContainer.id = 'reportContainer';
        document.getElementById('resultArea').appendChild(reportContainer);
    }
    reportContainer.innerHTML = '';

    // 1. Balance de Comprobación
    reportContainer.innerHTML += createTableHTML('Balance de Comprobación', data.trialBalance);

    // 2. Estado de Resultados
    reportContainer.innerHTML += createTableHTML('Estado de Resultados (Cuentas Nominales)', data.incomeStatement);

    // 3. Estado de Situación Financiera
    reportContainer.innerHTML += createTableHTML('Estado de Situación Financiera (Cuentas Reales)', data.financialPosition);
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

      setStatus('Calculando balances y estados financieros...', 'info');
      const balanceData = computeBalance(diaryRows, ledgerRows);
      lastBalance = balanceData;
      if (!balanceData || !balanceData.trialBalance || !balanceData.trialBalance.length) {
        setStatus('No se encontraron cuentas tras el cálculo. Compruebe los encabezados de columnas (Account / Cuenta / Cuenta contable).', 'error');
        exportBtn.disabled = true;
        saveBtn.disabled = true;
        return;
      }
      renderResult(balanceData); 
      exportBtn.disabled = false;
      setStatus(`Estados financieros generados correctamente.`, 'success');

      // Validar si el balance de comprobación está cuadrado
      const balanceCheck = checkBalanced(balanceData.trialBalance);
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

    const wb = XLSX.utils.book_new();

    const createSheet = (rows, sheetName) => {
        const data = rows.map(row => ({
          'Cuenta': row.Account,
          'Debe': row.Debit,
          'Haber': row.Credit,
          'Saldo': row.Balance
        }));

        const ws = XLSX.utils.json_to_sheet(data);
        const wscols = [{ wch: 50 }, { wch: 15 }, { wch: 15 }, { wch: 15 }];
        ws['!cols'] = wscols;

        // Formatos
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
          for (let C = 1; C <= 3; ++C) {
            const cell_address = { c: C, r: R };
            const cell_ref = XLSX.utils.encode_cell(cell_address);
            if (ws[cell_ref]) {
                ws[cell_ref].z = '#,##0.00';
                ws[cell_ref].t = 'n';
            }
          }
        }
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    };

    createSheet(lastBalance.trialBalance, 'Balance Comprobación');
    createSheet(lastBalance.incomeStatement, 'Estado Resultados');
    createSheet(lastBalance.financialPosition, 'Situación Financiera');

    // Detectar si estamos en Electron
    const isElectron = typeof require !== 'undefined' && typeof process !== 'undefined' && process.versions && process.versions.electron;

    if (isElectron) {
      try {
        const { dialog } = require('electron').remote || require('@electron/remote');
        const path = require('path');
        const os = require('os');

        const result = await dialog.showSaveDialog({
          title: 'Guardar Estados Financieros',
          defaultPath: path.join(os.homedir(), 'Downloads', 'estados_financieros.xlsx'),
          filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
        });

        if (!result.canceled && result.filePath) {
          XLSX.writeFile(wb, result.filePath);
          alert('Archivo guardado exitosamente en: ' + result.filePath);
        }
      } catch (err) {
        console.error('Error usando diálogo de Electron:', err);
        const path = require('path');
        const os = require('os');
        const filePath = path.join(os.homedir(), 'Downloads', 'estados_financieros.xlsx');
        XLSX.writeFile(wb, filePath);
        alert('Archivo guardado en: ' + filePath);
      }
    } else {
      XLSX.writeFile(wb, 'estados_financieros.xlsx');
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
        balance: lastBalance.trialBalance, // Guardar solo el trialBalance o todo? guardemos solo trialBalance por compatibilidad o todo
        fullReport: lastBalance // Guardar todo por si acaso
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
