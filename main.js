const { app, BrowserWindow } = require('electron');
const path = require('path');

// Inicializar @electron/remote
let remote;
try {
    remote = require('@electron/remote/main');
    remote.initialize();
} catch (err) {
    console.warn('@electron/remote no disponible:', err);
}

function createWindow() {
    const win = new BrowserWindow({
        width: 1200,
        height: 800,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false,
            enableRemoteModule: true // Habilitar para usar dialog
        },
        icon: path.join(__dirname, 'logo-usm.jpg')
    });

    // Habilitar remote para esta ventana
    if (remote) {
        remote.enable(win.webContents);
    }

    // Load the index.html from the web folder
    win.loadFile('web/index.html');

    // Remove default menu for a cleaner look
    win.setMenuBarVisibility(false);
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});
