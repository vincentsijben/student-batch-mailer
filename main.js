const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const { execFileSync } = require('child_process');
const fs = require('fs');
const ExcelJS = require('exceljs');

let mainWindow;
let logFilePath;
let templatesPath;
let userDataDir;
let uploadCacheDir;
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
let preparedAppleScriptPath = null;

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

function prepareOutlookAppleScript() {
  const bundledScriptPath = path.join(__dirname, 'outlook.scpt');
  if (!fs.existsSync(bundledScriptPath)) {
    return null;
  }
  if (preparedAppleScriptPath && fs.existsSync(preparedAppleScriptPath)) {
    return preparedAppleScriptPath;
  }
  try {
    const tempDir = path.join(app.getPath('temp'), 'student-batch-mailer');
    fs.mkdirSync(tempDir, { recursive: true });
    const tempPath = path.join(tempDir, 'outlook.scpt');
    const scriptContents = fs.readFileSync(bundledScriptPath, 'utf8');
    fs.writeFileSync(tempPath, scriptContents, 'utf8');
    preparedAppleScriptPath = tempPath;
    return preparedAppleScriptPath;
  } catch (error) {
    console.error('Unable to prepare Outlook AppleScript', error);
    return null;
  }
}

function appendLogEntries(entries) {
  if (!entries.length || !logFilePath) {
    return;
  }
  const current = readLogEntries();
  const updated = current.concat(entries);
  fs.writeFileSync(logFilePath, JSON.stringify(updated, null, 2), 'utf8');
}

function resetUploadCacheDir() {
  if (!uploadCacheDir) {
    return;
  }
  try {
    if (fs.existsSync(uploadCacheDir)) {
      fs.rmSync(uploadCacheDir, { recursive: true, force: true });
    }
  } catch (_) {
    // ignore cleanup errors
  }
  fs.mkdirSync(uploadCacheDir, { recursive: true });
}

app.whenReady().then(() => {
  userDataDir = app.getPath('userData');
  logFilePath = path.join(userDataDir, 'sent-log.json');
  templatesPath = path.join(userDataDir, 'templates.json');
  uploadCacheDir = path.join(userDataDir, 'upload-cache');
  ensureLogFile();
  ensureTemplatesFile();
  resetUploadCacheDir();
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

function normalizeCellValue(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((segment) => segment.text || '').join('');
    }
    if (typeof value.text !== 'undefined') {
      return String(value.text);
    }
    if (typeof value.result !== 'undefined') {
      return normalizeCellValue(value.result);
    }
    if (typeof value.hyperlink !== 'undefined' && typeof value.address === 'undefined') {
      return String(value.hyperlink);
    }
  }
  return String(value);
}

ipcMain.handle('parse-excel', async (_event, arrayBuffer) => {
  try {
    const buffer = Buffer.from(arrayBuffer);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    const worksheet = workbook.worksheets[0];
    if (!worksheet) {
      throw new Error('Workbook does not contain any sheets.');
    }

    const sheetValues = worksheet.getSheetValues().slice(1);
    const normalizedSheet = sheetValues
      .map((row) => {
        if (!row) {
          return null;
        }
        if (Array.isArray(row)) {
          return row.slice(1).map((cell) => normalizeCellValue(cell));
        }
        if (typeof row === 'object') {
          const cells = [];
          Object.keys(row).forEach((key) => {
            const index = Number(key);
            if (Number.isNaN(index) || index === 0) {
              return;
            }
            cells[index - 1] = normalizeCellValue(row[key]);
          });
          return cells;
        }
        return null;
      })
      .filter((row) => Array.isArray(row) && row.some((cell) => (cell || '').toString().trim() !== ''));

    if (!normalizedSheet.length) {
      throw new Error('Worksheet does not contain any data.');
    }

    const headerRow = normalizedSheet.shift();
    const headerMeta = headerRow.map((header) => {
      const text = normalizeCellValue(header).trim();
      const lower = text.toLowerCase();
      if (!text) {
        return { key: null, language: null };
      }
      if (['voornaam'].includes(lower)) {
        return { key: 'firstname', language: 'dutch' };
      }
      if (['achternaam'].includes(lower)) {
        return { key: 'lastname', language: 'dutch' };
      }
      if (['firstname', 'first name'].includes(lower)) {
        return { key: 'firstname', language: 'english' };
      }
      if (['lastname', 'last name'].includes(lower)) {
        return { key: 'lastname', language: 'english' };
      }
      if (['email', 'email address'].includes(lower)) {
        return { key: 'email', language: null };
      }
      if (['studentid', 'student id', 'id'].includes(lower)) {
        return { key: 'studentid', language: null };
      }
      return { key: null, language: null };
    });

    const hasDutchHeaders = headerMeta.some((meta) => meta.language === 'dutch');
    const records = normalizedSheet.map((row) => {
      const record = {
        firstname: '',
        lastname: '',
        email: '',
        studentid: '',
        sourceLanguage: hasDutchHeaders ? 'dutch' : 'english'
      };
      headerMeta.forEach((meta, index) => {
        if (!meta.key || !(meta.key in record)) {
          return;
        }
        const value = normalizeCellValue(row[index]).trim();
        if (value) {
          record[meta.key] = value;
        }
      });
      const lowerValues = [record.firstname, record.lastname, record.email].map((value) => value.toLowerCase());
      const headerKeywords = ['voornaam', 'achternaam', 'email', 'firstname', 'lastname', 'first name', 'last name'];
      if (lowerValues.some((value) => headerKeywords.includes(value))) {
        return null;
      }
      return record;
    }).filter(Boolean);

    const requiredColumns = ['firstname', 'lastname', 'email'];
    const missingColumns = requiredColumns.filter((column) =>
      records.every((row) => !row[column])
    );
    if (missingColumns.length) {
      throw new Error(`Missing required column(s): ${missingColumns.join(', ')}`);
    }
    const normalized = records.filter((student) => student.firstname || student.lastname || student.email);
    return { success: true, students: normalized };
  } catch (error) {
    return { success: false, message: error.message };
  }
});

ipcMain.handle('send-emails', async (_event, payload) => {
  const { matches } = payload;
  const scriptPath = prepareOutlookAppleScript();
  if (!scriptPath) {
    return { success: false, message: 'Outlook AppleScript not available.', hasLogEntries: hasLogEntries() };
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

ipcMain.handle('expand-paths', async (_event, targetPaths = []) => {
  try {
    if (!Array.isArray(targetPaths)) {
      throw new Error('Invalid path list.');
    }
    const results = [];
    const seen = new Set();
    function walk(entryPath) {
      if (!entryPath) {
        return;
      }
      const resolved = path.resolve(entryPath);
      if (seen.has(resolved)) {
        return;
      }
      seen.add(resolved);
      if (!fs.existsSync(resolved)) {
        return;
      }
      const stats = fs.statSync(resolved);
      if (stats.isDirectory()) {
        const children = fs.readdirSync(resolved);
        children.forEach((child) => walk(path.join(resolved, child)));
        return;
      }
      if (stats.isFile()) {
        results.push({
          path: resolved,
          name: path.basename(resolved)
        });
      }
    }
    targetPaths.forEach(walk);
    return {
      success: true,
      files: results.sort((a, b) => a.name.localeCompare(b.name))
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
});

ipcMain.handle('cache-uploaded-files', async (_event, files = []) => {
  try {
    if (!uploadCacheDir) {
      throw new Error('Upload cache not initialized.');
    }
    const ensureBuffer = (data) => {
      if (!data) {
        return null;
      }
      if (Buffer.isBuffer(data)) {
        return data;
      }
      if (data instanceof Uint8Array) {
        return Buffer.from(data);
      }
      if (data instanceof ArrayBuffer) {
        return Buffer.from(new Uint8Array(data));
      }
      return Buffer.from(data);
    };
    const cachedFiles = [];
    (files || []).forEach((file) => {
      const buffer = ensureBuffer(file && file.data);
      if (!buffer || !buffer.length) {
        return;
      }
      const safeName = (file && file.name ? path.basename(file.name) : 'file')
        .replace(/[<>:"/\\|?*\x00-\x1F]/g, '_')
        .slice(-200) || 'file';
      const uniqueDir = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
      const destinationDir = path.join(uploadCacheDir, uniqueDir);
      fs.mkdirSync(destinationDir, { recursive: true });
      const destination = path.join(destinationDir, safeName);
      fs.writeFileSync(destination, buffer);
      cachedFiles.push({
        path: destination,
        name: file && file.name ? file.name : safeName
      });
    });
    return { success: true, files: cachedFiles };
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
