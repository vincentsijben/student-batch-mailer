const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const { execFileSync } = require('child_process');
const fs = require('fs');
const XLSX = require('xlsx');

let mainWindow;
let logFilePath;
let templatesPath;
let userDataDir;
const pendingDebugMessages = [];
const originalConsole = {
  log: console.log,
  info: console.info,
  warn: console.warn,
  error: console.error
};
const consoleLevels = ['log', 'info', 'warn', 'error'];
consoleLevels.forEach((level) => {
  console[level] = (...args) => {
    try {
      originalConsole[level](...args);
    } catch (_) {
      // ignore logging failures
    }
    queueDebugMessage(level, 'main', args);
  };
});
const dutchDateFormatter = new Intl.DateTimeFormat('nl-NL', {
  timeZone: 'Europe/Amsterdam',
  year: 'numeric',
  month: '2-digit',
  day: '2-digit',
  hour: '2-digit',
  minute: '2-digit'
});

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1000,
    height: 700,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      sandbox: false
    }
  });

  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));

  mainWindow.webContents.on('did-finish-load', () => {
    flushPendingDebug();
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

function ensureLogFile() {
  if (!logFilePath) {
    return;
  }
  const dir = path.dirname(logFilePath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  if (!fs.existsSync(logFilePath)) {
    fs.writeFileSync(logFilePath, '[]', 'utf8');
  }
}

function ensureTemplatesFile() {
  if (!templatesPath) {
    return;
  }
  const dir = path.dirname(templatesPath);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  if (!fs.existsSync(templatesPath)) {
    fs.writeFileSync(templatesPath, '[]', 'utf8');
  }
}

function readLogEntries() {
  if (!logFilePath) {
    return [];
  }
  try {
    const raw = fs.readFileSync(logFilePath, 'utf8');
    return JSON.parse(raw);
  } catch {
    return [];
  }
}

function hasLogEntries() {
  return readLogEntries().length > 0;
}

function readTemplates() {
  if (!templatesPath) {
    return [];
  }
  try {
    const raw = fs.readFileSync(templatesPath, 'utf8');
    return JSON.parse(raw);
  } catch {
    return [];
  }
}

function writeTemplates(templates) {
  if (!templatesPath) {
    return;
  }
  fs.writeFileSync(templatesPath, JSON.stringify(templates, null, 2), 'utf8');
}

function queueDebugMessage(level, origin, message) {
  const payload = {
    level,
    origin,
    message: Array.isArray(message) ? formatDebugArgs(message) : (typeof message === 'string' ? message : formatDebugArgs([message]))
  };
  if (mainWindow && mainWindow.webContents && !mainWindow.webContents.isDestroyed()) {
    mainWindow.webContents.send('debug-message', payload);
  } else {
    pendingDebugMessages.push(payload);
  }
}

function flushPendingDebug() {
  if (!mainWindow || !mainWindow.webContents || mainWindow.webContents.isDestroyed()) {
    return;
  }
  while (pendingDebugMessages.length) {
    const payload = pendingDebugMessages.shift();
    mainWindow.webContents.send('debug-message', payload);
  }
}

function formatDebugArgs(values = []) {
  return values.map((value) => {
    if (typeof value === 'string') {
      return value;
    }
    try {
      return JSON.stringify(value);
    } catch (_) {
      return String(value);
    }
  }).join(' ');
}

function formatEmailBody(value) {
  const escaped = escapeHtml(value || '');
  return escaped
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .map((line) => line || '&nbsp;')
    .join('<br>');
}

function escapeHtml(text) {
  return text
    .toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function appendLogEntries(entries) {
  if (!entries.length || !logFilePath) {
    return;
  }
  const current = readLogEntries();
  const updated = current.concat(entries);
  fs.writeFileSync(logFilePath, JSON.stringify(updated, null, 2), 'utf8');
}

app.whenReady().then(() => {
  userDataDir = app.getPath('userData');
  logFilePath = path.join(userDataDir, 'sent-log.json');
  templatesPath = path.join(userDataDir, 'templates.json');
  ensureLogFile();
  ensureTemplatesFile();
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

ipcMain.handle('parse-excel', async (_event, arrayBuffer) => {
  try {
    const buffer = Buffer.from(arrayBuffer);
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    if (!sheetName) {
      throw new Error('Workbook does not contain any sheets.');
    }
    const worksheet = workbook.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false });
    const requiredColumns = ['firstname', 'lastname', 'email'];
    const normalizedRows = rows.map((row) => {
      const hasDutchNames = Boolean(row.voornaam || row.achternaam);
      const record = {
        firstname: String(row.firstname || row.firstName || row.voornaam || '').trim(),
        lastname: String(row.lastname || row.lastName || row.achternaam || '').trim(),
        email: String(row.email || '').trim(),
        studentid: String(row.studentid || row.studentId || '').trim(),
        sourceLanguage: hasDutchNames ? 'dutch' : 'english'
      };
      const lowerValues = [record.firstname, record.lastname, record.email].map((value) => value.toLowerCase());
      const headers = ['voornaam', 'achternaam', 'email', 'firstname', 'lastname', 'first name', 'last name'];
      if (lowerValues.some((value) => headers.includes(value))) {
        return null;
      }
      return record;
    }).filter(Boolean);
    const missingColumns = requiredColumns.filter((column) =>
      normalizedRows.every((row) => !row[column])
    );
    if (missingColumns.length) {
      throw new Error(`Missing required column(s): ${missingColumns.join(', ')}`);
    }
    const normalized = normalizedRows.filter(student => student.firstname || student.lastname || student.email);
    return { success: true, students: normalized };
  } catch (error) {
    return { success: false, message: error.message };
  }
});

ipcMain.handle('send-emails', async (_event, payload) => {
  const { matches } = payload;
  const scriptPath = path.join(__dirname, 'outlook.scpt');
  if (!fs.existsSync(scriptPath)) {
    return { success: false, message: 'Outlook AppleScript not found.', hasLogEntries: hasLogEntries() };
  }

  try {
    const logEntries = [];
    matches.forEach((match) => {
      if (!match.filePath || !fs.existsSync(match.filePath)) {
        throw new Error(`Attachment missing for ${match.firstname} ${match.lastname}`);
      }

      const subject = match.subject || 'Student Feedback';
      const emailBody = formatEmailBody(match.body) || 'Please see the attached feedback.';
      const recipientName = `${(match.firstname || '').trim()} ${(match.lastname || '').trim()}`.trim() || match.email;
      const attachmentArg = match.filePath || '';
      execFileSync('osascript', [
        scriptPath,
        attachmentArg,
        subject,
        emailBody,
        recipientName,
        match.email
      ], {
        stdio: 'pipe'
      });
      logEntries.push({
        timestamp: new Date().toISOString(),
        firstname: match.firstname,
        lastname: match.lastname,
        email: match.email,
        studentid: match.studentid,
        fileName: path.basename(match.filePath)
      });
    });
    appendLogEntries(logEntries);
    return { success: true, hasLogEntries: hasLogEntries() };
  } catch (error) {
    dialog.showErrorBox('Email Error', error.message);
    return { success: false, message: error.message, hasLogEntries: hasLogEntries() };
  }
});

ipcMain.handle('export-log', async () => {
  const entries = readLogEntries();
  if (!entries.length) {
    return { success: false, message: 'No sent emails to export yet.', hasLogEntries: false };
  }
  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow || undefined, {
    title: 'Export Sent Emails',
    defaultPath: 'student-feedback-sent.txt',
    filters: [{ name: 'Text Files', extensions: ['txt'] }]
  });
  if (canceled || !filePath) {
    return { success: false, message: 'Export canceled.', hasLogEntries: true };
  }
  const lines = entries.map((entry) => {
    const formattedDate = dutchDateFormatter.format(new Date(entry.timestamp));
    return `[${formattedDate}] - sent email to ${entry.firstname} ${entry.lastname} <${entry.email}> with attachment ${entry.fileName}`;
  });
  fs.writeFileSync(filePath, lines.join('\n'), 'utf8');
  return { success: true, filePath, hasLogEntries: true };
});

ipcMain.handle('get-log-status', async () => {
  return { success: true, hasEntries: hasLogEntries() };
});

ipcMain.handle('open-user-data', async () => {
  if (!userDataDir) {
    userDataDir = app.getPath('userData');
  }
  try {
    await shell.openPath(userDataDir);
    return { success: true, path: userDataDir };
  } catch (error) {
    return { success: false, message: error.message };
  }
});

ipcMain.handle('clear-log', async () => {
  if (!logFilePath) {
    return { success: false, message: 'Log path not initialized.', hasEntries: hasLogEntries() };
  }
  try {
    fs.writeFileSync(logFilePath, '[]', 'utf8');
    return { success: true, hasEntries: false };
  } catch (error) {
    return { success: false, message: error.message, hasEntries: hasLogEntries() };
  }
});

ipcMain.handle('list-templates', async () => {
  const templates = readTemplates();
  return { success: true, templates };
});

ipcMain.handle('save-template', async (_event, template) => {
  try {
    if (!template || !template.name) {
      return { success: false, message: 'Template name is required.', templates: readTemplates() };
    }
    const templates = readTemplates().filter((entry) => entry.name !== template.name);
    templates.push({
      name: template.name,
      subject: template.subject || '',
      body: template.body || ''
    });
    writeTemplates(templates);
    return { success: true, templates };
  } catch (error) {
    return { success: false, message: error.message, templates: readTemplates() };
  }
});

ipcMain.handle('delete-template', async (_event, name) => {
  try {
    if (!name) {
      return { success: false, message: 'Template name missing.', templates: readTemplates() };
    }
    const templates = readTemplates().filter((entry) => entry.name !== name);
    writeTemplates(templates);
    return { success: true, templates };
  } catch (error) {
    return { success: false, message: error.message, templates: readTemplates() };
  }
});
