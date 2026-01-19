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
    // Nuevo: Botón de reset
    const resetLicenseBtn = document.getElementById('resetLicense');

    if (!input || !btn || !error) {
      console.error('Elementos del modal de licencia no encontrados');
      return;
    }

    // Listener para resetear licencia
    if (resetLicenseBtn) {
      resetLicenseBtn.addEventListener('click', () => {
        if (confirm('¿Seguro que deseas eliminar la licencia guardada y reiniciar?')) {
          removeLicense(); // Función de license.js
          window.location.reload();
        }
      });
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
        // Feedback explícito por si el usuario no ve el texto
        alert('Código de licencia inválido.\n\nAsegúrate de usar uno de los códigos de la lista de licencias validas.');
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
      initTheme();
    });
  } else {
    checkLicenseOnStartup();
    setupLicenseModal();
    initTheme();
  }

  // ========================================
  // DARK MODE
  // ========================================
  function initTheme() {
    const toggleBtn = document.getElementById('themeToggle');
    const savedTheme = localStorage.getItem('theme');

    // Aplicar tema guardado o preferencia de sistema
    if (savedTheme === 'dark' || (!savedTheme && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
      document.body.classList.add('dark-mode');
      if (toggleBtn) toggleBtn.textContent = '☀️';
    }

    if (toggleBtn) {
      toggleBtn.addEventListener('click', () => {
        document.body.classList.toggle('dark-mode');
        const isDark = document.body.classList.contains('dark-mode');

        // Actualizar icono
        toggleBtn.textContent = isDark ? '☀️' : '🌙';

        // Guardar preferencia
        localStorage.setItem('theme', isDark ? 'dark' : 'light');

        // Actualizar gráficos si existen para aplicar nuevos colores
        if (window.lastBalanceData) {
          renderCharts(window.lastBalanceData.kpis, window.lastBalanceData.balance);
        }
      });
    }
  }

  // Nuevos inputs separados
  const inputDiary = document.getElementById('inputDiary');
  const inputLedger = document.getElementById('inputLedger');
  const computeBtn = document.getElementById('computeBtn');
  const exportBtn = document.getElementById('exportBtn');
  // const saveBtn = document.getElementById('saveBtn'); // Eliminado
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

          // VALIDACIÓN DE ENCABEZADOS:
          // A veces la primera fila es un título (ej. "LIBRO MAYOR") y XLSX la usa como header,
          // produciendo JSON basura y saltándose el fallback inteligente.
          // Comprobamos si las keys del JSON devuelto contienen palabras clave esperadas.
          if (json && json.length > 0) {
            const firstRowKeys = Object.keys(json[0]).map(k => String(k).toLowerCase());
            const expected = ['fecha', 'date', 'codigo', 'code', 'cuenta', 'account', 'detalle', 'debe', 'debit', 'haber', 'credit'];
            const hasValidHeaders = firstRowKeys.some(k => expected.some(exp => k.includes(exp)));

            if (!hasValidHeaders) {
              console.log('Encabezados automáticos no parecen contables. Forzando búsqueda manual...');
              json = []; // Forzar fallback
            }
          }

          // Fallback: si no devuelve filas o son invalidas, intentar analizar como matriz de matrices y construir objetos
          if ((!json || json.length === 0)) {
            const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
            // Smarter Header Detection
            // Instead of just taking the first non-empty row (which fails if there is a title like 'LIBRO MAYOR'),
            // we look for known keywords to identify the real header row.
            let headerRowIndex = -1;
            let maxScore = 0;

            const keywords = ['fecha', 'date', 'dia', 'código', 'codigo', 'code', 'cuenta', 'account', 'descripción', 'descripcion', 'detalle', 'debe', 'haber', 'debit', 'credit', 'saldo', 'balance'];

            // Scan first 15 rows (increased range)
            for (let i = 0; i < Math.min(15, aoa.length); i++) {
              const row = aoa[i];
              if (!row) continue;

              let score = 0;
              let hasContent = false;

              row.forEach(cell => {
                if (cell !== null && String(cell).trim() !== '') {
                  hasContent = true;
                  const cellStr = String(cell).toLowerCase().trim();
                  if (keywords.some(k => cellStr.includes(k))) {
                    score++;
                  }
                }
              });

              // If we found a row with more keywords, it's likely the header
              if (score > maxScore) {
                maxScore = score;
                headerRowIndex = i;
              } else if (headerRowIndex === -1 && hasContent && maxScore === 0) {
                // Fallback: preserve first non-empty row if no keywords found yet
                headerRowIndex = i;
              }
            }

            if (headerRowIndex >= 0) {
              console.log('Fila de encabezado detectada en índice:', headerRowIndex);


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

          // Estandarizar filas
          json = standardizeRows(json);
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

  // Normalizar claves de fila (eliminar espacios en nombres de columnas, trim, pero mantener case para display)
  // La comparación real se hace en getFirstValue de forma insensible
  function normalizeRowKeys(obj) {
    const out = {};
    for (const k of Object.keys(obj)) {
      out[String(k).trim()] = obj[k];
    }
    return out;
  }

  // Helper para buscar valores ignorando mayúsculas/minúsculas y acentos
  function getFirstValue(row, keys) {
    const rowKeys = Object.keys(row);

    for (const k of keys) {
      // 1. Busqueda Exacta
      if (row[k] !== undefined && row[k] !== null && row[k] !== '') return row[k];

      // 2. Busqueda Insensible (Case + Acentos)
      const kNorm = String(k).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

      const foundKey = rowKeys.find(rk => {
        const rkNorm = String(rk).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        return rkNorm === kNorm;
      });

      if (foundKey && row[foundKey] !== undefined && row[foundKey] !== null && row[foundKey] !== '') {
        return row[foundKey];
      }
    }
    return null;
  }

  // Estandarizar filas a estructura fija de 5 columnas
  // Estructura: Date, Code, Account, Debit, Credit
  function standardizeRows(rows) {
    if (!rows || rows.length === 0) return [];

    // Detectar encabezados de la primera fila
    const firstObj = normalizeRowKeys(rows[0]);
    // const keys = Object.keys(firstObj); // Ya no es estricto

    return rows.map(r => {
      const nr = normalizeRowKeys(r);
      const out = {
        Date: getFirstValue(nr, ['Fecha', 'Date', 'Day', 'Dia']),
        Code: getFirstValue(nr, ['Código', 'Codigo', 'Code', 'Cuenta Numero', 'Numero Cuenta']),
        Account: getFirstValue(nr, ['Cuenta', 'Account', 'Nombre Cuenta', 'Descripcion', 'Detalle']),
        Debit: getFirstValue(nr, ['Debe', 'Debit', 'Débito', 'Debito']),
        Credit: getFirstValue(nr, ['Haber', 'Credit', 'Crédito', 'Credito'])
      };

      // Si no encontró por nombre, intentar por posición (Indice 0, 1, 2, 3, 4)
      const vals = Object.values(nr);
      // Solo aplicar heurística posicional si faltan datos críticos (Cuenta o montos)
      if (!out.Account && !out.Debit && !out.Credit && vals.length >= 3) {
        // Asumiendo orden solicitado: Fecha?, Codigo, Nombre, Debe, Haber
        // A veces Fecha no está en Libro Mayor.

        // Estrategia: Si hay 5 columnas
        if (vals.length >= 5) {
          out.Date = vals[0];
          out.Code = vals[1];
          out.Account = vals[2];
          out.Debit = vals[3];
          out.Credit = vals[4];
        } else if (vals.length === 4) {
          // Asumir: Codigo, Nombre, Debe, Haber (Sin Fecha)
          out.Code = vals[0];
          out.Account = vals[1];
          out.Debit = vals[2];
          out.Credit = vals[3];
        } else if (vals.length === 3) {
          // Asumir: Nombre, Debe, Haber (Lo más básico)
          out.Account = vals[0];
          out.Debit = vals[1];
          out.Credit = vals[2];
        }
      }

      return out;
    });
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

    // Solo mostrar si es el balance final calculado
    if (balanceCheck.isFinalBalance) {
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
  }

  // ========================================
  // VALIDACIÓN DE LIBROS (NUEVO)
  // ========================================

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

    if (!diaryRows?.length && !ledgerRows?.length) {
      addMsg('No se cargaron archivos para validar.', 'warning');
      resultsDiv.className = 'validation-results warning';
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
    // Modificar preview para que se vea bonito el objeto estandarizado
    elem.textContent = JSON.stringify(rows.slice(0, 5), null, 2);
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
        // Asumimos que 'standardizeRows' ya corrió y tenemos keys: Date, Code, Account, Debit, Credit
        let account = r.Account;
        const code = r.Code;

        // Si no hay cuenta, ignorar
        if (!account && !code) return;

        // Construir clave única
        let key = account;
        if (code) {
          const c = String(code).trim();
          if (!account) key = c + ' - (Sin Nombre)';
          else if (!String(account).includes(c)) key = c + ' - ' + String(account).trim();
        }

        const finalKey = String(key || '').toUpperCase();

        // Sumar
        const prev = map.get(finalKey) || { Debit: 0, Credit: 0 };
        prev.Debit += Number(r.Debit) || 0;
        prev.Credit += Number(r.Credit) || 0;
        map.set(finalKey, prev);
      });
    };

    // Priorización: Si hay Libro Mayor, usarlo como fuente principal. 
    // Si no, usar Libro Diario. Esto evita duplicar montos si ambos contienen lo mismo.
    if (ledgerRows && ledgerRows.length > 0) {
      addRows(ledgerRows);
    } else if (diaryRows && diaryRows.length > 0) {
      addRows(diaryRows);
    }

    // Clasificación Estricta de cuentas
    const classifyAccount = (name, codeStr) => {
      const n = name.trim().toLowerCase();
      // Intentar extraer código de la clave si existe (ej "101 - CAJA")
      const match = n.match(/^(\d+)/);
      let c = codeStr || (match ? match[1] : '');

      // 1. Clasificación por Código (Si existe y es estándar)
      if (c) {
        const firstDigit = parseInt(c[0]);
        if (firstDigit >= 1 && firstDigit <= 3) return 'REAL'; // 1-Activo, 2-Pasivo, 3-Patrimonio
        if (firstDigit >= 4) return 'NOMINAL'; // 4-Ingresos, 5-Gastos, 6-Costos, 7-Costos Prod
      }

      // 2. Clasificación por Nombre (Palabras Clave)

      // EXCEPCIÓN CRÍTICA: Inventario
      if (n.includes('inventario') || n.includes('mercaderia') || n.includes('existencia')) {
        // Solo es nominal si dice explícitamente Costo
        if (n.includes('costo') && !n.includes('por cobrar')) return 'NOMINAL';
        return 'REAL'; // Activo
      }

      // Cuentas Nominales (Resultados)
      if (
        n.includes('venta') ||
        n.includes('ingreso') ||
        n.includes('gasto') ||
        n.includes('compra') ||
        n.includes('costo') ||
        n.includes('sueldo') ||
        n.includes('salario') ||
        n.includes('servicio') ||
        n.includes('honorario') ||
        n.includes('depreciacion') ||
        n.includes('amortizacion') ||
        n.includes('intereses pagados') ||
        n.includes('flete')
      ) {
        // Excepciones a palabras clave de gasto
        if (n.includes('por pagar') || n.includes('pagado por anticipado')) return 'REAL'; // Pasivo o Activo
        return 'NOMINAL';
      }

      // Por defecto REAL (Balance General)
      return 'REAL';
    };

    const trialBalance = [];
    const nominalAccounts = []; // Estado de Resultados

    // Categorías de Balance - Estructura agrupada
    const balanceSheet = {
      assets: {
        current: [],
        nonCurrent: []
      },
      liabilities: {
        current: [],
        nonCurrent: []
      },
      equity: []
    };

    let incomeStatementDebits = 0;
    let incomeStatementCredits = 0;

    for (const [accountKey, vals] of map.entries()) {
      const bal = vals.Debit - vals.Credit;

      // Determinación más precisa del nombre y código para lógica
      // accountKey suele ser "COD - NOMBRE" o solo "NOMBRE"
      const separatorIdx = accountKey.indexOf(' - ');
      let code = '';
      let name = accountKey;
      if (separatorIdx > -1) {
        code = accountKey.substring(0, separatorIdx);
        name = accountKey.substring(separatorIdx + 3);
      }

      const row = { Account: accountKey, Debit: vals.Debit, Credit: vals.Credit, Balance: bal, Code: code, CleanName: name };

      // 1. Balance de Comprobación (Todos)
      trialBalance.push(row);

      // 2. Clasificación
      const type = classifyAccount(name, code);

      if (type === 'NOMINAL') {
        nominalAccounts.push(row);
        incomeStatementDebits += vals.Debit;
        incomeStatementCredits += vals.Credit;
      } else {
        // Es REAL (Balance General)
        // Sub-clasificación: Activo, Pasivo, Patrimonio

        const n = name.toLowerCase();
        let group = 'ASSET'; // Default

        // Detectar Pasivos y Patrimonio
        if (code && (code.startsWith('2'))) group = 'LIABILITY';
        else if (code && (code.startsWith('3'))) group = 'EQUITY';
        else {
          // Heurística por nombre
          if (n.includes('capital') || n.includes('utilidad') || n.includes('superavit') || n.includes('reserva') || n.includes('patrimonio') || n.includes('resultado acumulado')) group = 'EQUITY';
          else if (n.includes('pasivo') || n.includes('por pagar') || n.includes('proveedor') || n.includes('obligacion') || n.includes('anticipo de clientes')) group = 'LIABILITY';
          else group = 'ASSET';
        }

        // Lógica Especial: Sobregiro Bancario
        if (group === 'ASSET' && (n.includes('banco') || n.includes('cuenta corriente'))) {
          // Si el saldo es Negativo (Acreedor), es un Pasivo (Sobregiro)
          if (row.Balance < 0) {
            group = 'LIABILITY';
            row.Account = row.Account + ' (Sobregiro)';
            row.Balance = Math.abs(row.Balance); // Mostrar como positivo en la sección de pasivo
            const temp = row.Debit; row.Debit = row.Credit; row.Credit = temp;
          }
        }

        // Asignar a grupos Corriente / No Corriente
        if (group === 'ASSET') {
          // Corriente: Caja, Bancos, Clientes, Inventario, IVA Credito
          if (
            n.includes('caja') || n.includes('banco') || n.includes('cliente') || n.includes('por cobrar') ||
            n.includes('inventario') || n.includes('mercaderia') || n.includes('iva') || n.includes('anticipo') ||
            (code && code.startsWith('11')) // Plan contable común: 11 es corriente
          ) {
            balanceSheet.assets.current.push(row);
          } else {
            // No Corriente: Propiedad Planta Equipo, Muebles, Vehiculos, Software, Depreciacion Acumulada
            balanceSheet.assets.nonCurrent.push(row);
          }
        } else if (group === 'LIABILITY') {
          // Corriente: Corto plazo, Impuestos, Proveedores
          // No Corriente: Largo plazo, Prestamos
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

    // Ordenar Listas
    const sortFn = (a, b) => a.Account.localeCompare(b.Account);
    trialBalance.sort(sortFn);
    nominalAccounts.sort(sortFn);
    balanceSheet.assets.current.sort(sortFn);
    balanceSheet.assets.nonCurrent.sort(sortFn);
    balanceSheet.liabilities.current.sort(sortFn);
    balanceSheet.liabilities.nonCurrent.sort(sortFn);
    balanceSheet.equity.sort(sortFn);

    // Calcular Utilidad/Pérdida
    const netIncome = incomeStatementCredits - incomeStatementDebits;

    // Agregar Utilidad al Patrimonio
    const netIncomeRow = {
      Account: 'UTILIDAD (PÉRDIDA) DEL EJERCICIO',
      Debit: netIncome < 0 ? Math.abs(netIncome) : 0,
      Credit: netIncome > 0 ? netIncome : 0,
      Balance: -netIncome // Para que coincida con lógica de Patrimonio (Acreedor = Negativo)
    };

    balanceSheet.equity.push(netIncomeRow);

    // ===================================
    // ANÁLISIS VERTICAL (Preparación)
    // ===================================

    // Total Activos
    let totalAssets = 0;
    [...balanceSheet.assets.current, ...balanceSheet.assets.nonCurrent].forEach(r => totalAssets += Number(r.Balance));

    // Total Ingresos (Base para Est. Resultados)
    let totalRevenue = nominalAccounts.reduce((acc, r) => {
      if (r.Balance < 0) return acc + Math.abs(Number(r.Balance));
      return acc;
    }, 0);

    // Asignar Porcentajes
    const computePerc = (list, base) => {
      list.forEach(r => {
        const val = Math.abs(Number(r.Balance));
        r.Percentage = base ? (val / base) : 0;
      });
    };

    computePerc([...balanceSheet.assets.current, ...balanceSheet.assets.nonCurrent], totalAssets);
    computePerc(nominalAccounts, totalRevenue);

    // Retorno Estructurado
    const financialPositionFlat = [
      ...balanceSheet.assets.current,
      ...balanceSheet.assets.nonCurrent,
      ...balanceSheet.liabilities.current,
      ...balanceSheet.liabilities.nonCurrent,
      ...balanceSheet.equity
    ];

    return {
      trialBalance,
      incomeStatement: nominalAccounts,
      financialPosition: financialPositionFlat, // Usar el aplanado para la vista web actual
      balanceSheetStructured: balanceSheet, // Nuevo objeto estructurado para exportación experta
      netIncome,
      isDiary: (diaryRows && diaryRows.length > 0)
    };
  }

  // Helper para renderizar una tabla específica
  function createTableHTML(title, rows, showTotals = true, showPercentage = false) {
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
              ${showPercentage ? '<th>% A.V.</th>' : ''}
            </tr>
          </thead>
          <tbody>
    `;

    rows.forEach(r => {
      let balClass = '';
      if (Number(r.Balance) > 0) balClass = 'balance-positive';
      else if (Number(r.Balance) < 0) balClass = 'balance-negative';
      else balClass = 'balance-zero';

      let percStr = '';
      if (showPercentage && r.Percentage !== undefined) {
        percStr = (r.Percentage * 100).toFixed(2) + '%';
      }

      html += `
        <tr>
          <td>${r.Account}</td>
          <td>${formatNumber(r.Debit)}</td>
          <td>${formatNumber(r.Credit)}</td>
          <td class="${balClass}">${formatNumber(r.Balance)}</td>
          ${showPercentage ? `<td class="text-right">${percStr}</td>` : ''}
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
               ${showPercentage ? '<td></td>' : ''}
            </tr>
          </tfoot>
      `;
    }

    html += `</table></div>`;
    return html;
  }

  // ========================================
  // DASHBOARD & GRÁFICOS
  // ========================================

  let charts = {}; // Guardar instancias de Chart.js para destruirlas al actualizar

  // Global Error Handler to catch crash issues
  window.addEventListener('error', function (event) {
    console.error('Global Error:', event.error);
    alert('Ha ocurrido un error inesperado:\n' + (event.error ? event.error.message : event.message));
  });

  function calculateKPIs(data) {
    // Usar la estructura robusta si está disponible
    if (data.balanceSheetStructured) {
      const s = data.balanceSheetStructured;

      const currentAssets = s.assets.current.reduce((a, b) => a + Number(b.Balance), 0);
      const nonCurrentAssets = s.assets.nonCurrent.reduce((a, b) => a + Number(b.Balance), 0);
      const totalAssets = currentAssets + nonCurrentAssets;

      // Pasivos suelen ser negativos en Balance, usamos Math.abs para KPI
      const currentLiabilities = s.liabilities.current.reduce((a, b) => a + Math.abs(Number(b.Balance)), 0);
      const nonCurrentLiabilities = s.liabilities.nonCurrent.reduce((a, b) => a + Math.abs(Number(b.Balance)), 0);
      const totalLiabilities = currentLiabilities + nonCurrentLiabilities;

      const totalEquity = s.equity.reduce((a, b) => a + Math.abs(Number(b.Balance)), 0);

      // Ingresos (Nominales Haber)
      // Iteramos incomeStatement para sumar Ingresos (Balance < 0)
      const totalRevenue = data.incomeStatement.reduce((acc, r) => {
        if (r.Balance < 0) return acc + Math.abs(Number(r.Balance));
        return acc;
      }, 0);

      const netIncome = data.netIncome;

      return {
        liquidity: currentLiabilities ? (currentAssets / currentLiabilities) : 0,
        debtRatio: totalAssets ? (totalLiabilities / totalAssets) : 0,
        netMargin: totalRevenue ? (netIncome / totalRevenue) : 0,
        netResult: netIncome,
        totalAssets,
        totalLiabilities,
        totalEquity,
        totalRevenue,
        totalExpenses: totalRevenue - netIncome
      };
    }

    // Fallback Legacy (Si por alguna razón falla la estructura)
    const { financialPosition, netIncome } = data;
    let totalAssets = 0;
    let totalLiabilities = 0;
    let totalEquity = 0;

    financialPosition.forEach(r => {
      const bal = Number(r.Balance);
      // Heurística simple
      const match = r.Account.trim().match(/^(\d+)/);
      if (match) {
        const digit = parseInt(match[1][0]);
        if (digit === 1) totalAssets += bal;
        else if (digit === 2) totalLiabilities += Math.abs(bal);
        else if (digit === 3) totalEquity += Math.abs(bal);
      } else {
        if (bal > 0) totalAssets += bal;
        else {
          const n = r.Account.toLowerCase();
          if (n.includes('capital') || n.includes('utilidad') || n.includes('patrimonio')) totalEquity += Math.abs(bal);
          else totalLiabilities += Math.abs(bal);
        }
      }
    });

    // Estimación corriente
    let currentAssets = totalAssets * 0.6; // No podemos saber sin estructura
    let currentLiabilities = totalLiabilities * 0.5;

    // Ingresos
    const totalRevenue = data.incomeStatement.reduce((acc, r) => {
      if (r.Balance < 0) return acc + Math.abs(Number(r.Balance));
      return acc;
    }, 0);

    return {
      liquidity: currentLiabilities ? (currentAssets / currentLiabilities) : 0,
      debtRatio: totalAssets ? (totalLiabilities / totalAssets) : 0,
      netMargin: totalRevenue ? (netIncome / totalRevenue) : 0,
      netResult: netIncome,
      totalAssets,
      totalLiabilities,
      totalEquity,
      totalRevenue,
      totalExpenses: totalRevenue - netIncome
    };
  }
  // Charts instances
  let chartInstances = {};


  function renderCharts(kpis, balance) {
    if (typeof Chart === 'undefined') return;

    // Detectar si estamos en dark mode para colores de texto
    const isDark = document.body.classList.contains('dark-mode');
    const textColor = isDark ? '#94a3b8' : '#666';
    const gridColor = isDark ? '#334155' : '#ddd';

    const ctxAssets = document.getElementById('chartAssets').getContext('2d');
    const ctxResults = document.getElementById('chartIncome').getContext('2d');

    // Destruir anteriores si existen
    if (chartInstances.assets) chartInstances.assets.destroy();
    if (chartInstances.results) chartInstances.results.destroy();

    // 1. Gráfico Ecuación Patrimonial (Doughnut)
    chartInstances.assets = new Chart(ctxAssets, {
      type: 'doughnut',
      data: {
        labels: ['Pasivo', 'Patrimonio'],
        datasets: [{
          data: [kpis.totalLiabilities, kpis.totalEquity],
          backgroundColor: ['#e81500', '#00ac69'],
          borderWidth: 0
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            labels: { color: textColor }
          }
        }
      }
    });

    // 2. Gráfico Resultados (Bar)
    chartInstances.results = new Chart(ctxResults, {
      type: 'bar',
      data: {
        labels: ['Ingresos', 'Gastos', 'Utilidad'],
        datasets: [{
          label: 'Montos',
          data: [kpis.totalRevenue, kpis.totalExpenses, kpis.netResult],
          backgroundColor: ['#0061f2', '#f4a100', kpis.netResult >= 0 ? '#00ac69' : '#e81500'],
          borderRadius: 4
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          title: { display: false }
        },
        scales: {
          y: {
            beginAtZero: true,
            grid: { color: gridColor },
            ticks: { color: textColor }
          },
          x: {
            grid: { display: false },
            ticks: { color: textColor }
          }
        }
      }
    });
  }

  function updateDashboard(data) {
    const dashboard = document.getElementById('financialDashboard');
    if (!dashboard) return;

    const kpis = calculateKPIs(data);

    // Actualizar Textos
    document.getElementById('kpiLiquidity').textContent = kpis.liquidity.toFixed(2);
    document.getElementById('kpiDebt').textContent = (kpis.debtRatio * 100).toFixed(1) + '%';
    document.getElementById('kpiMargin').textContent = (kpis.netMargin * 100).toFixed(1) + '%';
    document.getElementById('kpiNetResult').textContent = formatNumber(kpis.netResult);

    // Mostrar dashboard
    dashboard.hidden = false;
    dashboard.style.display = 'block';

    // Renderizar gráficos
    // Pequeño delay para asegurar visibilidad
    setTimeout(() => renderCharts(kpis), 100);
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
    reportContainer.innerHTML += createTableHTML('Balance de Comprobación', data.trialBalance, true, false);

    // 2. Estado de Resultados (con % sobre Ventas)
    reportContainer.innerHTML += createTableHTML('Estado de Resultados (Cuentas Nominales)', data.incomeStatement, true, true);

    // 3. Estado de Situación Financiera (con % sobre Activo Total)
    reportContainer.innerHTML += createTableHTML('Estado de Situación Financiera (Cuentas Reales)', data.financialPosition, true, true);
  }

  computeBtn.addEventListener('click', async () => {

    console.log('Click en boton generar');

    const fileDiary = inputDiary.files[0];
    const fileLedger = inputLedger.files[0];

    if (!fileDiary && !fileLedger) {
      alert('Seleccione al menos un archivo (Libro Diario o Libro Mayor)');
      return;
    }

    alert('Iniciando proceso de validación y cálculo...');

    try {
      computeBtn.disabled = true;
      setStatus('Leyendo archivos...', 'info');

      // Limpiar resultados anteriores
      if (document.getElementById('validationResults')) document.getElementById('validationResults').hidden = true;

      let diaryRows = [];
      let ledgerRows = [];

      // Leer Diario
      if (fileDiary) {
        try {
          diaryRows = await readExcelFile(fileDiary);
        } catch (e) {
          console.error('Error leyendo diario', e);
          setStatus('Error al leer Libro Diario', 'error');
          return;
        }
      }

      // Leer Mayor
      if (fileLedger) {
        try {
          ledgerRows = await readExcelFile(fileLedger);
        } catch (e) {
          console.error('Error leyendo mayor', e);
          setStatus('Error al leer Libro Mayor', 'error');
          return;
        }
      }

      if ((!diaryRows || !diaryRows.length) && (!ledgerRows || !ledgerRows.length)) {
        setStatus('No se encontraron datos en los archivos seleccionados.', 'error');
        exportBtn.disabled = true;
        return;
      }

      // Ejecutar Validador
      validateBooks(diaryRows, ledgerRows);

      // Calcular Balance General Unificado (Fusionando lo que haya)
      // Nota: Si hay ambos, computeBalance sumará ambos. 
      // Lo ideal para el balance general es usar los dos si son complementarios, 
      // O usar el Mayor si es un resumen.
      // Suposición: Si el usuario sube ambos, quiere validar consistencia, 
      // pero el Balance General sale del Mayor (o del Diario acumulado).
      // Si sumamos ambos, duplicamos.

      // Lógica de decisión de fuente para Balance General:
      // 1. Si hay Mayor, usar Mayor.
      // 2. Si no, usar Diario.
      // 3. (Caso borde) Si el usuario quiere mezclar, el sistema actual 'computeBalance' suma todo.
      //    Para evitar duplicidad si suben lo mismo:
      //    Vamos a usar solo UNO para el cálculo del Balance General final, priorizando el Mayor si existe y está cuadrado.

      // Decisión: Usaremos TODOS los datos disponibles para el 'computeBalance' original 
      // porque así estaba diseñado (sumar todos los inputs).
      // PERO, el usuario podría estar subiendo el mismo contenido dos veces (formato diario y formato mayor).
      // ADVERTENCIA: computeBalance suma las filas.

      // Cambiamos estrategia: pasar SOLO el que parezca más completo o ambos si así se desea?
      // El requerimiento dice: "Carga tu libro diario y libro mayor... Genera un balance"
      // Si sube ambos, ¿se duplican?
      // Hagamos que computeBalance use ambos, pero advirtamos o dejemos que el usuario decida.
      // O mejor: Si sube ambos, asumimos que son complementarios (ej. movimientos del mes y saldos anteriores)
      // O que son redundantes. Si son redundantes, el balance se duplicará.

      // FIX: Para el Balance General, usaremos la unión de ambos, asumiendo que el usuario sabe lo que hace.
      // Si se duplica, se verá en los montos.

      const balanceData = computeBalance(diaryRows, ledgerRows);

      // Flag para exportación
      balanceData.isDiary = (diaryRows && diaryRows.length > 0);

      lastBalance = balanceData;

      if (!balanceData || !balanceData.trialBalance || !balanceData.trialBalance.length) {
        setStatus('No se encontraron cuentas tras el cálculo.', 'error');
        exportBtn.disabled = true;
        return;
      }

      exportBtn.disabled = false;
      setStatus(`Estados financieros generados correctamente.`, 'success');

      // Validar si el balance de comprobación FINAL está cuadrado
      const balanceCheck = checkBalanced(balanceData.trialBalance);
      balanceCheck.isFinalBalance = true; // Para diferenciar en UI
      showBalanceStatus(balanceCheck);

      // Calcular y mostrar Dashboard
      updateDashboard(balanceData);

      // Mostrar tabla de resultados
      renderResult(balanceData);

    } catch (err) {
      console.error(err);
      setStatus('Error procesando: ' + (err.message || err), 'error');
      alert('Error procesando: ' + err.message);
    } finally {
      computeBtn.disabled = false;
    }
  });

  exportBtn.addEventListener('click', async () => {
    if (!lastBalance) return;

    // Función helper para añadir hoja con estilo profesional
    const addToSheet = (wb, name, title, blocks, options = {}) => {
      const wsData = [];

      // Configuración de columnas por defecto
      const headers = options.headers || ['Cuenta', 'Debe', 'Haber', 'Saldo', '% A.V.'];
      const mapRow = options.mapRow || ((r) => {
        let perc = '';
        if (r.Percentage !== undefined) perc = (r.Percentage * 100).toFixed(2) + '%';
        return [
          r.CleanName || r.Account,
          r.Debit,
          r.Credit,
          r.Balance,
          perc
        ];
      });

      // Título Reporte
      wsData.push([title]);
      wsData.push(['']); // Espacio

      const merges = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }]; // Merge título

      let currentRow = 2;

      blocks.forEach(block => {
        // Subtítulo de Sección (ej. ACTIVO CORRIENTE)
        if (block.title) {
          wsData.push([block.title]);
          // Merge subtítulo
          merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow, c: 4 } });
          currentRow++;
        }

        // Encabezados de tabla
        wsData.push(headers);
        currentRow++;

        // Filas de datos
        block.rows.forEach(r => {
          wsData.push(mapRow(r));
          currentRow++;
        });

        // Totales de bloque
        if (block.rows.length > 0) {
          const sumBal = block.rows.reduce((a, b) => a + Number(b.Balance), 0);
          // Ajuste básico para totales: Asumimos que el saldo está en la columna index 3 (D) o 4 (E) aprox.
          // Para "Libro Mayor" con Código, los índices cambian. 
          // Haremos algo genérico: Poner "Total" en col 0 y el saldo en una columna final si coincide.

          // Por simplicidad en esta versión flexible, solo ponemos Total Saldo alineado a la derecha
          const totalRow = new Array(headers.length).fill('');
          totalRow[0] = 'Total ' + block.title;

          // Buscar columna de saldo por nombre
          const balIdx = headers.findIndex(h => h.toLowerCase().includes('saldo'));
          if (balIdx > -1) totalRow[balIdx] = sumBal;
          else totalRow[headers.length - 1] = sumBal; // Fallback al final

          wsData.push(totalRow);
          currentRow++;
          wsData.push(['']); // Espacio tras bloque
          currentRow++;
        } else {
          const emptyRow = new Array(headers.length).fill('');
          emptyRow[0] = '(Sin cuentas)';
          wsData.push(emptyRow);
          currentRow++;
          wsData.push(['']);
          currentRow++;
        }
      });

      // Crear hoja
      const ws = XLSX.utils.aoa_to_sheet(wsData);

      // Aplicar merges
      ws['!merges'] = merges;

      // Anchos (Estimación básica)
      const wscols = headers.map((h, i) => {
        if (i === 0) return { wch: 50 }; // Nombre/Código ancho
        return { wch: 15 }; // Montos
      });
      ws['!cols'] = wscols;

      // Formato de celdas (Iterar para aplicar estilo numérico)
      const range = XLSX.utils.decode_range(ws['!ref']);
      for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = 0; C <= headers.length; ++C) {
          const cellRef = XLSX.utils.encode_cell({ c: C, r: R });
          if (ws[cellRef] && typeof ws[cellRef].v === 'number') {
            ws[cellRef].z = '#,##0.00';
            ws[cellRef].t = 'n';
          }
        }
      }

      XLSX.utils.book_append_sheet(wb, ws, name);
    };

    // Preparar Bloques para Situación Financiera
    const struct = lastBalance.balanceSheetStructured;

    // Calcular totales para validación
    const sumAssets = [...struct.assets.current, ...struct.assets.nonCurrent].reduce((a, b) => a + Number(b.Balance), 0);
    const sumLiab = [...struct.liabilities.current, ...struct.liabilities.nonCurrent].reduce((a, b) => a + Math.abs(Number(b.Balance)), 0);
    const sumEquity = struct.equity.reduce((a, b) => a + Math.abs(Number(b.Balance)), 0);
    // Nota: Liabilities y Equity suelen venir con signo negativo en nuestra lógica, usamos Math.abs para la ecuación A = P + E

    const wb = XLSX.utils.book_new();

    // 1. Hoja: Estado de Situación Financiera
    addToSheet(wb, 'Sit. Financiera', 'ESTADO DE SITUACIÓN FINANCIERA', [
      { title: 'ACTIVO CORRIENTE', rows: struct.assets.current },
      { title: 'ACTIVO NO CORRIENTE', rows: struct.assets.nonCurrent },
      { title: 'PASIVO CORRIENTE', rows: struct.liabilities.current },
      { title: 'PASIVO NO CORRIENTE', rows: struct.liabilities.nonCurrent },
      { title: 'PATRIMONIO', rows: struct.equity }
    ]);

    // Añadir validación al final de la hoja 1 manualmente
    const ws1 = wb.Sheets['Sit. Financiera'];
    // Encontrar ultima fila
    const range1 = XLSX.utils.decode_range(ws1['!ref']);
    let lastRow = range1.e.r + 2;

    XLSX.utils.sheet_add_aoa(ws1, [
      ['ECUACIÓN PATRIMONIAL'],
      ['Total Activo', sumAssets],
      ['Total Pasivo + Patrimonio', sumLiab + sumEquity],
      ['Diferencia', sumAssets - (sumLiab + sumEquity)],
      [(sumAssets - (sumLiab + sumEquity) < 0.01) ? 'VALIDADO OK' : 'DESCUADRADO']
    ], { origin: -1 });

    // 2. Hoja: Estado de Resultados
    addToSheet(wb, 'Est. Resultados', 'ESTADO DE RESULTADOS', [
      { title: 'CUENTAS DE RESULTADO', rows: lastBalance.incomeStatement }
    ]);

    // Añadir resumen final utilidad
    const ws2 = wb.Sheets['Est. Resultados'];
    XLSX.utils.sheet_add_aoa(ws2, [
      [''],
      ['RESULTADO DEL EJERCICIO', lastBalance.netIncome]
    ], { origin: -1 });


    // 3. Hoja: Libro Mayor (Formato Especifico: Código, Cuenta, Debe, Haber)
    // Se genera SIEMPRE que hay datos, usando trialBalance (que es la lista unificada)
    addToSheet(wb, 'Libro Mayor', 'LIBRO MAYOR',
      [{ title: 'MOVIMIENTOS POR CUENTA', rows: lastBalance.trialBalance }],
      {
        headers: ['Código', 'Cuenta', 'Debe', 'Haber', 'Saldo'],
        mapRow: (r) => [
          r.Code || '', // Código primero
          r.CleanName || r.Account,
          r.Debit,
          r.Credit,
          r.Balance
        ]
      }
    );

    // 4. Hoja: Balance de Comprobación (Formato Standard con %)
    addToSheet(wb, 'B. Comprobación', 'BALANCE DE COMPROBACIÓN', [
      { title: 'TODAS LAS CUENTAS', rows: lastBalance.trialBalance }
    ]);

    // Guardar
    // Guardar
    if (typeof require !== 'undefined' && typeof require('path') !== 'undefined') {
      // Entorno Electron
      try {
        const remote = require('@electron/remote');
        const { dialog } = remote;

        const { filePath } = await dialog.showSaveDialog({
          title: 'Guardar Estados Financieros',
          defaultPath: 'Estados_Financieros_Profesional.xlsx',
          filters: [{ name: 'Excel', extensions: ['xlsx'] }]
        });

        if (filePath) {
          XLSX.writeFile(wb, filePath);
          setStatus(`Archivo guardado exitosamente en: ${filePath}`, 'success');
          alert('Archivo exportado correctamente.');
        }

      } catch (e) {
        console.error('Error usando dialog de Electron:', e);
        // Fallback por si falla remote
        XLSX.writeFile(wb, 'Estados_Financieros_Profesional.xlsx');
        alert('Se intentó guardar en la carpeta de la aplicación.');
      }
    } else {
      // Entorno Web normal
      XLSX.writeFile(wb, 'Estados_Financieros_Profesional.xlsx');
    }

  });

})();
