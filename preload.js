const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  parseExcel: (arrayBuffer) => ipcRenderer.invoke('parse-excel', arrayBuffer),
  sendEmails: (payload) => ipcRenderer.invoke('send-emails', payload),
  getLogStatus: () => ipcRenderer.invoke('get-log-status'),
  openSentLogs: () => ipcRenderer.invoke('open-sent-logs'),
  listTemplates: () => ipcRenderer.invoke('list-templates'),
  saveTemplate: (template) => ipcRenderer.invoke('save-template', template),
  deleteTemplate: (name) => ipcRenderer.invoke('delete-template', name),
  // Copies files/folders into the session cache; the originals are never modified.
  importPaths: (paths, skipOriginals) => ipcRenderer.invoke('import-paths', paths, skipOriginals),
  discardFiles: (copyPaths) => ipcRenderer.invoke('discard-files', copyPaths),
  cacheUploadedFiles: (files) => ipcRenderer.invoke('cache-uploaded-files', files),
  // File.path was removed in Electron 32; this is the supported replacement.
  // Returns '' for files that have no location on disk.
  getPathForFile: (file) => {
    try {
      return webUtils.getPathForFile(file) || '';
    } catch (_) {
      return '';
    }
  },
  onSendProgress: (callback) => ipcRenderer.on('send-progress', (_event, progress) => callback(progress)),
  onDebugMessage: (callback) => ipcRenderer.on('debug-message', (_event, payload) => callback(payload))
});
