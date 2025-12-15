const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  parseExcel: (arrayBuffer) => ipcRenderer.invoke('parse-excel', arrayBuffer),
  sendEmails: (payload) => ipcRenderer.invoke('send-emails', payload),
  exportLog: () => ipcRenderer.invoke('export-log'),
  getLogStatus: () => ipcRenderer.invoke('get-log-status'),
  openUserData: () => ipcRenderer.invoke('open-user-data'),
  clearLog: () => ipcRenderer.invoke('clear-log'),
  listTemplates: () => ipcRenderer.invoke('list-templates'),
  saveTemplate: (template) => ipcRenderer.invoke('save-template', template),
  deleteTemplate: (name) => ipcRenderer.invoke('delete-template', name),
  onDebugMessage: (callback) => ipcRenderer.on('debug-message', callback)
});
