// Script para eliminar la licencia guardada
// Ejecuta este archivo con: node eliminar_licencia.js

console.log('🗑️  Eliminando licencia guardada...\n');

// Simular localStorage para Node.js
const { app } = require('electron');
const path = require('path');
const fs = require('fs');

// La ruta donde Electron guarda localStorage
const userDataPath = app.getPath('userData');
const localStoragePath = path.join(userDataPath, 'Local Storage', 'leveldb');

console.log('📁 Ubicación de datos de usuario:', userDataPath);
console.log('📁 Ubicación de localStorage:', localStoragePath);

console.log('\n⚠️  NOTA: Este script debe ejecutarse desde Electron.');
console.log('Para eliminar la licencia manualmente:');
console.log('1. Abre la aplicación Electron');
console.log('2. Presiona F12 para abrir DevTools');
console.log('3. Ve a la pestaña "Console"');
console.log('4. Ejecuta: localStorage.removeItem("app_license_key")');
console.log('5. Recarga la aplicación (Ctrl+R o Cmd+R)');
console.log('\n✅ La licencia será eliminada y verás el modal de activación nuevamente.');
