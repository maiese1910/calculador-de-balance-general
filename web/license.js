// license.js - Sistema de Licencias
// Funciones para validar y gestionar licencias

/**
 * Calcula el checksum MD5 simplificado para validación de licencia
 * Usa un algoritmo simple compatible con el generador Python
 */
function simpleHash(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        const char = str.charCodeAt(i);
        hash = ((hash << 5) - hash) + char;
        hash = hash & hash;
    }
    // Convertir a hex y tomar primeros 4 caracteres
    const hex = Math.abs(hash).toString(16).toUpperCase();
    return (hex + '0000').substring(0, 4);
}

/**
 * Valida un código de licencia
 * Formato: XXXX-XXXX-XXXX-XXXX
 * Los primeros 3 bloques son datos, el 4to es checksum
 */
function validateLicenseKey(key) {
    if (!key || typeof key !== 'string') return false;

    // Limpiar espacios y convertir a mayúsculas
    key = key.trim().toUpperCase();

    // Verificar longitud
    if (key.length !== 19) return false;

    // Verificar formato
    const parts = key.split('-');
    if (parts.length !== 4) return false;

    // Verificar que cada parte tenga 4 caracteres
    for (const part of parts) {
        if (part.length !== 4) return false;
        // Verificar que solo contenga letras y números
        if (!/^[A-Z0-9]+$/.test(part)) return false;
    }

    // Lista de licencias válidas (generadas por generar_licencias.py)
    const validLicenses = [
        'CS23-3AL1-R9NO-234D',
        '2E80-7QTI-05S3-4A3D',
        '6DYE-VTU8-LJJ0-3BD2',
        'U44O-6QKB-IANN-4139',
        'B535-CDZ8-I30I-ED29',
        'H0NE-R2QC-SX5D-FE02',
        '85F3-4Q5C-UW84-0D7E',
        'EVTK-T17J-PUNN-0874',
        'Q3JZ-HTBZ-4E0U-BE2F',
        'GGB4-2805-16WS-0041'
    ];

    return validLicenses.includes(key);
}

/**
 * Guarda la licencia en localStorage
 */
function saveLicense(key) {
    try {
        localStorage.setItem('app_license_key', key);
        return true;
    } catch (err) {
        console.error('Error guardando licencia:', err);
        return false;
    }
}

/**
 * Obtiene la licencia guardada
 */
function getSavedLicense() {
    try {
        return localStorage.getItem('app_license_key');
    } catch (err) {
        console.error('Error leyendo licencia:', err);
        return null;
    }
}

/**
 * Verifica si hay una licencia válida guardada
 */
function hasValidLicense() {
    const saved = getSavedLicense();
    return saved && validateLicenseKey(saved);
}

/**
 * Elimina la licencia guardada (para testing)
 */
function removeLicense() {
    try {
        localStorage.removeItem('app_license_key');
        return true;
    } catch (err) {
        console.error('Error eliminando licencia:', err);
        return false;
    }
}

// Exportar funciones para uso en app.js
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        validateLicenseKey,
        saveLicense,
        getSavedLicense,
        hasValidLicense,
        removeLicense
    };
}
