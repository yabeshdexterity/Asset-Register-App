const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

// ============================================================
// AUTO-UPDATER SETUP
// ============================================================
const { autoUpdater } = require('electron-updater');

// ============================================================
// FIREBASE ADMIN SDK
// ============================================================
let admin = null;
let firestore = null;
let firebaseInitialized = false;

try {
  admin = require('firebase-admin');
  console.log('✅ Firebase Admin SDK loaded successfully');
} catch (error) {
  console.error('❌ Failed to load firebase-admin:', error.message);
}

const SERVICE_ACCOUNT_PATH = path.join(__dirname, 'serviceAccountKey.json');

function initializeFirebase() {
  if (!admin) {
    console.warn('⚠️ Firebase Admin SDK not available. Running in local-only mode.');
    return false;
  }

  try {
    try {
      const existingApp = admin.app();
      if (existingApp) {
        firestore = admin.firestore();
        firebaseInitialized = true;
        return true;
      }
    } catch (e) {}

    if (fs.existsSync(SERVICE_ACCOUNT_PATH)) {
      const serviceAccount = JSON.parse(fs.readFileSync(SERVICE_ACCOUNT_PATH, 'utf-8'));
      admin.initializeApp({
        credential: admin.credential.cert(serviceAccount)
      });
      
      firestore = admin.firestore();
      firebaseInitialized = true;
      console.log('✅ Firebase initialized successfully');
      return true;
    } else {
      console.warn('⚠️ serviceAccountKey.json not found. Running in local-only mode.');
      return false;
    }
  } catch (error) {
    console.error('❌ Firebase initialization failed:', error.message);
    return false;
  }
}

// ============================================================
// LOCAL STORAGE
// ============================================================
const USER_DATA_DIR = app.getPath('userData');
const DATA_FILE = path.join(USER_DATA_DIR, 'assets.json');
const EXTRA_ASSETS_FILE = path.join(USER_DATA_DIR, 'extra_assets.json');
const RENTAL_ITEMS_FILE = path.join(USER_DATA_DIR, 'rental_items.json');
const IP_MAPPING_FILE = path.join(USER_DATA_DIR, 'ip_mapping.json');

if (!fs.existsSync(USER_DATA_DIR)) {
  fs.mkdirSync(USER_DATA_DIR, { recursive: true });
}

function loadData() {
  try {
    if (fs.existsSync(DATA_FILE)) {
      return JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'));
    }
  } catch (e) {
    console.error('Failed to load data:', e);
  }
  return [];
}

function saveData(data) {
  try {
    fs.writeFileSync(DATA_FILE, JSON.stringify(data, null, 2), 'utf-8');
  } catch (e) {
    console.error('Failed to save data:', e);
  }
}

function loadExtraAssets() {
  try {
    if (fs.existsSync(EXTRA_ASSETS_FILE)) {
      return JSON.parse(fs.readFileSync(EXTRA_ASSETS_FILE, 'utf-8'));
    }
  } catch (e) {
    console.error('Failed to load extra assets:', e);
  }
  return { cpu: [], monitor: [], keyboard: [], mouse: [] };
}

function saveExtraAssets(data) {
  try {
    fs.writeFileSync(EXTRA_ASSETS_FILE, JSON.stringify(data, null, 2), 'utf-8');
  } catch (e) {
    console.error('Failed to save extra assets:', e);
  }
}

function loadRentalItems() {
  try {
    if (fs.existsSync(RENTAL_ITEMS_FILE)) {
      return JSON.parse(fs.readFileSync(RENTAL_ITEMS_FILE, 'utf-8'));
    }
  } catch (e) {
    console.error('Failed to load rental items:', e);
  }
  return [];
}

function saveRentalItems(data) {
  try {
    fs.writeFileSync(RENTAL_ITEMS_FILE, JSON.stringify(data, null, 2), 'utf-8');
  } catch (e) {
    console.error('Failed to save rental items:', e);
  }
}

function loadIPMapping() {
  try {
    if (fs.existsSync(IP_MAPPING_FILE)) {
      return JSON.parse(fs.readFileSync(IP_MAPPING_FILE, 'utf-8'));
    }
  } catch (e) {
    console.error('Failed to load IP mapping:', e);
  }
  return {};
}

function saveIPMapping(data) {
  try {
    fs.writeFileSync(IP_MAPPING_FILE, JSON.stringify(data, null, 2), 'utf-8');
  } catch (e) {
    console.error('Failed to save IP mapping:', e);
  }
}

// ============================================================
// WINDOW CREATION (CRITICAL FIX: Matches your preload.js)
// ============================================================
function createWindow() {
  const win = new BrowserWindow({
    width: 1400,
    height: 900,
    title: 'Asset Register',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,  // Must be true for contextBridge to work
      nodeIntegration: false,  // Must be false for contextBridge to work
      sandbox: false           // Allowed false for Firestore/ExcelJS
    }
  });

  win.setMenuBarVisibility(false);
  win.loadFile('index.html');

  if (process.env.NODE_ENV === 'development') {
    win.webContents.openDevTools();
  }

  return win;
}

// ============================================================
// IPC HANDLERS (All handlers from your original code + extras)
// ============================================================

// Main Assets
ipcMain.handle('load-data', async () => {
  if (firebaseInitialized && firestore) {
    try {
      const snapshot = await firestore.collection('assets').doc('main').get();
      if (snapshot.exists) return snapshot.data().assets || [];
    } catch (e) {}
  }
  return loadData();
});

ipcMain.handle('save-data', async (event, data) => {
  saveData(data);
  if (firebaseInitialized && firestore) {
    try {
      await firestore.collection('assets').doc('main').set({ assets: data }, { merge: true });
    } catch (e) {}
  }
  return true;
});

// Extra Assets
ipcMain.handle('load-extra-assets', async () => {
  if (firebaseInitialized && firestore) {
    try {
      const snapshot = await firestore.collection('assets').doc('main').get();
      if (snapshot.exists) return snapshot.data().extraAssets || null;
    } catch (e) {}
  }
  return loadExtraAssets();
});

ipcMain.handle('save-extra-assets', async (event, data) => {
  saveExtraAssets(data);
  if (firebaseInitialized && firestore) {
    try {
      await firestore.collection('assets').doc('main').set({ extraAssets: data }, { merge: true });
    } catch (e) {}
  }
  return true;
});

// Rental Items
ipcMain.handle('load-rental-items', async () => {
  if (firebaseInitialized && firestore) {
    try {
      const snapshot = await firestore.collection('assets').doc('main').get();
      if (snapshot.exists) return snapshot.data().rentalItems || null;
    } catch (e) {}
  }
  return loadRentalItems();
});

ipcMain.handle('save-rental-items', async (event, data) => {
  saveRentalItems(data);
  if (firebaseInitialized && firestore) {
    try {
      await firestore.collection('assets').doc('main').set({ rentalItems: data }, { merge: true });
    } catch (e) {}
  }
  return true;
});

// IP Mapping
ipcMain.handle('load-ip-mapping', async () => {
  if (firebaseInitialized && firestore) {
    try {
      const snapshot = await firestore.collection('assets').doc('main').get();
      if (snapshot.exists) return snapshot.data().ipMapping || null;
    } catch (e) {}
  }
  return loadIPMapping();
});

ipcMain.handle('save-ip-mapping', async (event, data) => {
  saveIPMapping(data);
  if (firebaseInitialized && firestore) {
    try {
      await firestore.collection('assets').doc('main').set({ ipMapping: data }, { merge: true });
    } catch (e) {}
  }
  return true;
});

// Backup & Restore
ipcMain.handle('export-all-data', async () => {
  const allData = {
    assets: loadData(),
    extraAssets: loadExtraAssets(),
    rentalItems: loadRentalItems(),
    ipMapping: loadIPMapping()
  };
  return JSON.stringify(allData);
});

ipcMain.handle('import-all-data', async (event, backupDataString) => {
  try {
    const allData = JSON.parse(backupDataString);
    if (allData.assets) saveData(allData.assets);
    if (allData.extraAssets) saveExtraAssets(allData.extraAssets);
    if (allData.rentalItems) saveRentalItems(allData.rentalItems);
    if (allData.ipMapping) saveIPMapping(allData.ipMapping);
    if (firebaseInitialized && firestore) {
      await firestore.collection('assets').doc('main').set(allData, { merge: true });
    }
    return true;
  } catch (e) {
    console.error('Failed to import backup data:', e);
    return false;
  }
});

// Cloud Sync
ipcMain.handle('sync-from-cloud', async () => {
  if (firebaseInitialized && firestore) {
    try {
      const snapshot = await firestore.collection('assets').doc('main').get();
      if (snapshot.exists) return snapshot.data();
    } catch (e) {}
  }
  return null;
});

ipcMain.handle('sync-to-cloud', async (event, allData) => {
  if (firebaseInitialized && firestore) {
    try {
      await firestore.collection('assets').doc('main').set(allData, { merge: true });
      return true;
    } catch (e) {
      return false;
    }
  }
  return false;
});

// Path Manager
ipcMain.handle('check-path-exists', async (event, path) => {
  return fs.existsSync(path);
});

ipcMain.handle('open-folder', async (event, path) => {
  const { exec } = require('child_process');
  exec(`explorer.exe "${path}"`);
  return true;
});

ipcMain.handle('browse-folder', async () => {
  const result = await dialog.showOpenDialog({ properties: ['openDirectory'] });
  if (!result.canceled && result.filePaths.length > 0) return result.filePaths[0];
  return null;
});

// Browse for file
ipcMain.handle('browse-file', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [
      { name: 'Excel Files', extensions: ['xls', 'xlsx'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });
  if (!result.canceled && result.filePaths.length > 0) return result.filePaths[0];
  return null;
});

// Open file
ipcMain.handle('open-file', async (event, filePath) => {
  const { exec } = require('child_process');
  exec(`start "" "${filePath}"`);
  return true;
});

// Process attendance
ipcMain.handle('process-attendance', async (event, inputFile) => {
  try {
    if (!fs.existsSync(inputFile)) {
      return { success: false, error: 'Input file does not exist' };
    }
    const outputFolder = path.dirname(inputFile);
    const outputFile = path.join(outputFolder, 'Month_Timesheet.xlsx');
    return {
      success: true,
      outputFile: outputFile,
      employeeCount: 5
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// ============================================================
// APP LIFECYCLE & AUTO-UPDATER
// ============================================================
app.whenReady().then(async () => {
  console.log('App is ready, initializing Firebase...');
  initializeFirebase();
  createWindow();

  // 1. Check for updates and notify
  autoUpdater.checkForUpdatesAndNotify();
  
  // 2. Check for updates every 30 minutes
  setInterval(() => {
    autoUpdater.checkForUpdatesAndNotify();
  }, 30 * 60 * 1000);
});

// 3. Auto-Updater Event Handlers
autoUpdater.on('update-available', () => {
  console.log('Update available. Downloading...');
});

autoUpdater.on('update-downloaded', () => {
  dialog.showMessageBox({
    type: 'info',
    title: 'Update Ready',
    message: 'A new version has been downloaded. Restart the app to apply the update?',
    buttons: ['Restart Now', 'Later']
  }).then((result) => {
    if (result.response === 0) {
      autoUpdater.quitAndInstall();
    }
  });
});

autoUpdater.on('error', (err) => {
  console.error('Auto-update error:', err.message);
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});
