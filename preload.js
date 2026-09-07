const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('assetAPI', {
  loadData: () => ipcRenderer.invoke('load-data'),
  saveData: (data) => ipcRenderer.invoke('save-data', data),
  loadExtraAssets: () => ipcRenderer.invoke('load-extra-assets'),
  saveExtraAssets: (data) => ipcRenderer.invoke('save-extra-assets', data),
  loadRentalItems: () => ipcRenderer.invoke('load-rental-items'),
  saveRentalItems: (data) => ipcRenderer.invoke('save-rental-items', data),
  loadIPMapping: () => ipcRenderer.invoke('load-ip-mapping'),
  saveIPMapping: (data) => ipcRenderer.invoke('save-ip-mapping', data),
  checkPathExists: (path) => ipcRenderer.invoke('check-path-exists', path),
  openFolder: (path) => ipcRenderer.invoke('open-folder', path),
  browseFolder: () => ipcRenderer.invoke('browse-folder'),
  // Add to existing contextBridge
  browseFile: () => ipcRenderer.invoke('browse-file'),
  openFile: (path) => ipcRenderer.invoke('open-file', path),
  processAttendance: (filePath) => ipcRenderer.invoke('process-attendance', filePath),
  
  // Backup & Restore
  exportAllData: () => ipcRenderer.invoke('export-all-data'),
  importAllData: (data) => ipcRenderer.invoke('import-all-data', data),

  // Cloud Sync (Firestore only)
  syncFromCloud: () => ipcRenderer.invoke('sync-from-cloud'),
  syncToCloud: (data) => ipcRenderer.invoke('sync-to-cloud', data)
});

// Listen for sync events from main process
ipcRenderer.on('data-synced', (event, data) => {
  window.dispatchEvent(new CustomEvent('data-synced', { detail: data }));
});

