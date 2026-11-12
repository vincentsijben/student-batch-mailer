const { app, BrowserWindow, ipcMain, dialog, shell, Notification, powerSaveBlocker } = require('electron');
const path = require('path');
const { execFile } = require('child_process');
const fs = require('fs');
const ExcelJS = require('exceljs');

let mainWindow;
let sentLogsDir;
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
let sendInProgress = false;
let sendProgress = { done: 0, total: 0 };
// Microsoft 365 accepts about 30 emails per minute per mailbox; faster batches
// can be throttled or bounced. One email every 2.5 s stays safely below that.
const SEND_INTERVAL_MS = 2500;
const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const EMAIL_PATTERN = /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/;

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

  // Closing the window quits the app, which would stop a batch halfway.
  mainWindow.on('close', (event) => {
    if (!sendInProgress) {
      return;
    }
    const choice = dialog.showMessageBoxSync(mainWindow, {
      type: 'warning',
      buttons: ['Keep sending', 'Quit anyway'],
      defaultId: 0,
      cancelId: 0,
      message: 'Emails are still being sent',
      detail: `${sendProgress.done} of ${sendProgress.total} sent. If you quit now, the rest is not sent. The ones already sent are listed in the send report.`
    });
    if (choice === 0) {
      event.preventDefault();
    }
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
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

function hasSentLogs() {
  if (!sentLogsDir || !fs.existsSync(sentLogsDir)) {
    return false;
  }
  try {
    return fs.readdirSync(sentLogsDir).some((name) => name.endsWith('.html'));
  } catch {
    return false;
  }
}

function formatLogTimestampForFileName(date) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}_${pad(date.getHours())}-${pad(date.getMinutes())}-${pad(date.getSeconds())}`;
}

// Copies the report stylesheet into <sent-logs>/css/style.css so every report
// links to one shared file. Re-copied on each run so style updates propagate.
function ensureSentLogStylesheet() {
  if (!sentLogsDir) {
    return;
  }
  try {
    const source = path.join(__dirname, 'report', 'style.css');
    const cssDir = path.join(sentLogsDir, 'css');
    fs.mkdirSync(cssDir, { recursive: true });
    fs.copyFileSync(source, path.join(cssDir, 'style.css'));
  } catch (error) {
    console.error('Unable to copy report stylesheet', error);
  }
}

function studentLabel(record) {
  return `${record.firstname || ''} ${record.lastname || ''}`.trim() || record.email || '—';
}

function renderStudentTable(rows, { withFile = true, withTime = false } = {}) {
  if (!rows.length) {
    return '<p class="empty">None</p>';
  }
  const head = [
    withTime ? '<th>Time</th>' : '',
    '<th>Student</th>',
    '<th>Email</th>',
    withFile ? '<th>Attachment</th>' : ''
  ].join('');
  const body = rows.map((row) => [
    '<tr>',
    withTime ? `<td class="time">${escapeHtml(row.timestamp ? dutchDateFormatter.format(new Date(row.timestamp)) : '')}</td>` : '',
    `<td class="student">${escapeHtml(studentLabel(row))}${row.studentid ? ` <span class="muted">(${escapeHtml(row.studentid)})</span>` : ''}</td>`,
    `<td><a href="mailto:${escapeHtml(row.email || '')}">${escapeHtml(row.email || '')}</a></td>`,
    withFile ? `<td class="file">${escapeHtml(row.fileName || '')}</td>` : '',
    '</tr>'
  ].join('')).join('\n');
  return `<table><thead><tr>${head}</tr></thead><tbody>${body}</tbody></table>`;
}

// Writes one HTML report per send action into the sent-logs folder.
function writeSentLog(entries, report = {}, failure = null) {
  if (!sentLogsDir) {
    return null;
  }
  const skipped = Array.isArray(report.skipped) ? report.skipped : [];
  const unmatchedStudents = Array.isArray(report.unmatchedStudents) ? report.unmatchedStudents : [];
  const unmatchedFiles = Array.isArray(report.unmatchedFiles) ? report.unmatchedFiles : [];
  const problems = Array.isArray(report.problems) ? report.problems : [];
  if (!entries.length && !failure) {
    return null;
  }
  fs.mkdirSync(sentLogsDir, { recursive: true });
  ensureSentLogStylesheet();

  const now = new Date();
  const fileStamp = formatLogTimestampForFileName(now);
  const filePath = path.join(sentLogsDir, `sent-${fileStamp}.html`);
  const title = `Sent emails – ${dutchDateFormatter.format(now)}`;
  const plural = (n, word) => `${n} ${word}${n === 1 ? '' : 's'}`;

  const summaryItems = [
    `<li class="ok"><strong>${entries.length}</strong> sent</li>`,
    `<li class="${skipped.length ? 'warn' : ''}"><strong>${skipped.length}</strong> matched but unchecked</li>`,
    `<li class="${unmatchedStudents.length ? 'warn' : ''}"><strong>${unmatchedStudents.length}</strong> without a file</li>`,
    `<li class="${unmatchedFiles.length ? 'warn' : ''}"><strong>${unmatchedFiles.length}</strong> file${unmatchedFiles.length === 1 ? '' : 's'} without a student</li>`,
    `<li class="${problems.length ? 'warn' : ''}"><strong>${problems.length}</strong> open problem${problems.length === 1 ? '' : 's'}</li>`
  ].join('');

  const html = `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<link rel="stylesheet" href="css/style.css">
</head>
<body>
<main>
  <header>
    <h1>Student Batch Mailer – send report</h1>
    <p class="meta">${escapeHtml(dutchDateFormatter.format(now))}${report.subjectTemplate ? ` · Subject: <em>${escapeHtml(report.subjectTemplate)}</em>` : ''}</p>
    ${failure ? `<p class="failure">⚠ Sending stopped with an error: ${escapeHtml(failure)}. The emails listed under “Sent” went out before the error.</p>` : ''}
    <ul class="summary">${summaryItems}</ul>
  </header>

  ${problems.length ? `<section>
    <h2><span class="dot warn"></span>Open problems at the time of sending <span class="count">${plural(problems.length, 'item')}</span></h2>
    <ul class="files">${problems.map((text) => `<li>${escapeHtml(text)}</li>`).join('')}</ul>
  </section>` : ''}

  <section>
    <h2><span class="dot ok"></span>Sent <span class="count">${plural(entries.length, 'email')}</span></h2>
    ${renderStudentTable(entries, { withFile: true, withTime: true })}
  </section>

  <section>
    <h2><span class="dot warn"></span>Matched but not sent (unchecked) <span class="count">${plural(skipped.length, 'student')}</span></h2>
    ${renderStudentTable(skipped, { withFile: true })}
  </section>

  <section>
    <h2><span class="dot warn"></span>Students without a matching file <span class="count">${plural(unmatchedStudents.length, 'student')}</span></h2>
    ${renderStudentTable(unmatchedStudents, { withFile: false })}
  </section>

  <section>
    <h2><span class="dot muted"></span>Files without a matching student <span class="count">${plural(unmatchedFiles.length, 'file')}</span></h2>
    ${unmatchedFiles.length
      ? `<ul class="files">${unmatchedFiles.map((name) => `<li>${escapeHtml(name)}</li>`).join('')}</ul>`
      : '<p class="empty">None</p>'}
  </section>

  <footer>Generated by Student Batch Mailer · ${escapeHtml(path.basename(filePath))}</footer>
</main>
</body>
</html>
`;
  fs.writeFileSync(filePath, html, 'utf8');
  return filePath;
}

function readTemplates() {
  if (!templatesPath) {
    return [];
  }
  try {
    const parsed = JSON.parse(fs.readFileSync(templatesPath, 'utf8'));
    return Array.isArray(parsed) ? parsed : [];
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

// The script sets Outlook's *plain text* body, so the text must not contain
// HTML (that showed up literally as <br>&nbsp; in received emails).
function plainTextBody(value) {
  return (value || '').toString().replace(/\r\n?/g, '\n');
}

// Progress in the terminal (pnpm run start), next to the progress bar in the app.
function logSend(message) {
  originalConsole.log(`[${new Date().toLocaleTimeString('nl-NL')}] ${message}`);
}

function notifyDone(title, body) {
  try {
    if (Notification.isSupported() && !(mainWindow && mainWindow.isFocused())) {
      new Notification({ title, body, silent: false }).show();
    }
  } catch (_) {
    // notifications are a nice-to-have
  }
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
  sentLogsDir = path.join(userDataDir, 'sent-logs');
  templatesPath = path.join(userDataDir, 'templates.json');
  uploadCacheDir = path.join(userDataDir, 'upload-cache');
  fs.mkdirSync(sentLogsDir, { recursive: true });
  ensureSentLogStylesheet();
  ensureTemplatesFile();
  seedDefaultTemplates();
  resetUploadCacheDir();
  createWindow();
});

// Single-window utility: closing the window quits the app on every platform,
// instead of the macOS default of lingering in the Dock until Cmd+Q.
app.on('window-all-closed', () => {
  app.quit();
});

function normalizeCellValue(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }
  if (typeof value === 'object') {
    if (Array.isArray(value.richText)) {
      return value.richText.map((segment) => segment.text || '').join('');
    }
    if (typeof value.text !== 'undefined') {
      // Hyperlink cells: `text` can itself be a rich-text object.
      return normalizeCellValue(value.text);
    }
    if (typeof value.result !== 'undefined') {
      return normalizeCellValue(value.result);
    }
    if (typeof value.hyperlink !== 'undefined') {
      return String(value.hyperlink).replace(/^mailto:/i, '');
    }
    // Formula without a cached result, error values (#N/A), etc.
    return '';
  }
  return String(value);
}

const HEADER_ALIASES = [
  { key: 'firstname', language: 'dutch', names: ['voornaam', 'roepnaam'] },
  { key: 'lastname', language: 'dutch', names: ['achternaam'] },
  { key: 'firstname', language: 'english', names: ['firstname', 'first name', 'given name'] },
  { key: 'lastname', language: 'english', names: ['lastname', 'last name', 'surname', 'family name'] },
  { key: 'email', language: null, names: ['email', 'e-mail', 'email address', 'e-mail address', 'emailadres', 'e-mailadres', 'mail'] },
  { key: 'studentid', language: null, names: ['studentid', 'student id', 'id', 'studentnummer', 'student number'] }
];
const REQUIRED_COLUMNS = ['firstname', 'lastname', 'email'];

function classifyHeader(text) {
  const lower = (text || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const alias = lower && HEADER_ALIASES.find((entry) => entry.names.includes(lower));
  return alias ? { key: alias.key, language: alias.language } : { key: null, language: null };
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

    // Keep the real Excel row number so problems can point at "row 14".
    const rows = [];
    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const values = Array.isArray(row.values) ? row.values : [];
      const cells = [];
      for (let i = 1; i < values.length; i += 1) {
        cells.push(normalizeCellValue(values[i]).trim());
      }
      if (cells.some(Boolean)) {
        rows.push({ rowNumber, cells });
      }
    });

    if (!rows.length) {
      throw new Error('Worksheet does not contain any data.');
    }

    // The header is the first row (within the first 10) that has all required
    // columns, so a title row like "Klas 2A" above the headers doesn't break it.
    let headerIndex = rows.slice(0, 10).findIndex((row) => {
      const keys = row.cells.map((cell) => classifyHeader(cell).key);
      return REQUIRED_COLUMNS.every((key) => keys.includes(key));
    });
    if (headerIndex === -1) {
      headerIndex = 0;
    }
    const headerMeta = rows[headerIndex].cells.map(classifyHeader);
    const foundKeys = headerMeta.map((meta) => meta.key);
    const missingColumns = REQUIRED_COLUMNS.filter((key) => !foundKeys.includes(key));
    if (missingColumns.length) {
      throw new Error(`Missing required column(s): ${missingColumns.join(', ')}. Expected headers like firstname/voornaam, lastname/achternaam, email.`);
    }
    const duplicateKeys = REQUIRED_COLUMNS.filter((key) => foundKeys.filter((k) => k === key).length > 1);
    if (duplicateKeys.length) {
      throw new Error(`The roster has more than one column for: ${duplicateKeys.join(', ')}. Remove or rename the extra column.`);
    }

    const hasDutchHeaders = headerMeta.some((meta) => meta.language === 'dutch');
    const students = rows.slice(headerIndex + 1).map(({ rowNumber, cells }) => {
      const record = {
        firstname: '',
        lastname: '',
        email: '',
        studentid: '',
        rowNumber,
        sourceLanguage: hasDutchHeaders ? 'dutch' : 'english'
      };
      headerMeta.forEach((meta, index) => {
        if (meta.key && cells[index]) {
          record[meta.key] = cells[index];
        }
      });
      // Skip repeated header rows (e.g. when several class lists were pasted together).
      const repeatsHeader = [record.firstname, record.lastname, record.email]
        .some((value) => classifyHeader(value).key);
      if (repeatsHeader) {
        return null;
      }
      return record;
    }).filter((student) => student && (student.firstname || student.lastname || student.email));

    if (!students.length) {
      throw new Error('The roster has headers but no student rows.');
    }
    return { success: true, students };
  } catch (error) {
    return { success: false, message: error.message };
  }
});

function runOsascript(args) {
  return new Promise((resolve, reject) => {
    execFile('osascript', args, { timeout: 120000 }, (error, stdout, stderr) => {
      if (error) {
        // error.message repeats the full command line (incl. the email body);
        // stderr holds the actual AppleScript/Outlook error.
        const detail = (stderr || '').toString().trim() || (error.killed ? 'Outlook did not respond within 2 minutes.' : error.message);
        reject(new Error(detail));
        return;
      }
      resolve(stdout);
    });
  });
}

function safeWriteSentLog(...args) {
  try {
    return writeSentLog(...args);
  } catch (error) {
    console.error('Unable to write sent log', error);
    return null;
  }
}

ipcMain.handle('send-emails', async (event, payload) => {
  const { matches = [], report = {}, test = false } = payload || {};
  if (sendInProgress) {
    return { success: false, message: 'A send is already in progress.', sentIds: [] };
  }
  if (!matches.length) {
    return { success: false, message: 'Nothing to send.', sentIds: [] };
  }
  const scriptPath = prepareOutlookAppleScript();
  if (!scriptPath) {
    return { success: false, message: 'Outlook AppleScript not available.', sentIds: [] };
  }

  // Check every email before sending any, so a bad row can't stop the batch halfway.
  const recipientLabel = (match) => `${(match.firstname || '').trim()} ${(match.lastname || '').trim()}`.trim() || match.email || 'unknown student';
  const preflight = [];
  const seenEmails = new Map();
  matches.forEach((match) => {
    const who = recipientLabel(match);
    if (!match.filePath || !fs.existsSync(match.filePath)) {
      preflight.push(`attachment missing for ${who} (${match.fileName || 'no file'})`);
    }
    const email = (match.email || '').trim();
    if (!EMAIL_PATTERN.test(email)) {
      preflight.push(`invalid email address for ${who}: "${email}"`);
    } else if (!test && seenEmails.has(email.toLowerCase())) {
      preflight.push(`${who} and ${seenEmails.get(email.toLowerCase())} have the same email address (${email})`);
    } else {
      seenEmails.set(email.toLowerCase(), who);
    }
  });
  if (preflight.length) {
    return { success: false, message: `Nothing was sent: ${preflight.join('; ')}.`, sentIds: [] };
  }

  const notify = (progress) => {
    if (!event.sender.isDestroyed()) {
      event.sender.send('send-progress', progress);
    }
  };

  sendInProgress = true;
  sendProgress = { done: 0, total: matches.length };
  // Keep sending while the app is in the background: no App Nap, no idle sleep.
  // (Closing the lid still puts the Mac to sleep; sending resumes after waking.)
  const blockerId = powerSaveBlocker.start('prevent-app-suspension');
  const logEntries = [];
  const sentIds = [];
  const total = matches.length;
  const kind = test ? 'test email' : 'email';
  // Test sends are at most 5 emails, well under the limit: no pause needed.
  const interval = test ? 0 : SEND_INTERVAL_MS;
  const started = Date.now();
  logSend(`Sending ${total} ${kind}${total === 1 ? '' : 's'} via Outlook${interval && total > 1 ? ` (one every ${interval / 1000}s, about ${formatDuration((total - 1) * interval)})` : ''}…`);
  let lastSendStarted = 0;
  try {
    for (let index = 0; index < total; index += 1) {
      const match = matches[index];
      const recipientName = match.recipientName || recipientLabel(match);
      const wait = lastSendStarted ? interval - (Date.now() - lastSendStarted) : 0;
      if (wait > 0) {
        notify({ done: index, total, current: recipientName, test, waitMs: wait, remainingMs: wait + (total - index - 1) * interval });
        await sleep(wait);
      }
      lastSendStarted = Date.now();
      notify({ done: index, total, current: recipientName, test, remainingMs: (total - index - 1) * interval });
      await runOsascript([
        scriptPath,
        match.filePath,
        match.subject || 'Student Feedback',
        plainTextBody(match.body) || 'Please see the attached feedback.',
        recipientName,
        match.email.trim()
      ]);
      logSend(`  ${String(index + 1).padStart(String(total).length)}/${total} ✓ ${recipientName} <${match.email.trim()}> — ${path.basename(match.filePath)}`);
      sentIds.push(match.id);
      sendProgress.done = sentIds.length;
      logEntries.push({
        timestamp: new Date().toISOString(),
        firstname: match.firstname,
        lastname: match.lastname,
        email: match.email,
        studentid: match.studentid,
        fileName: path.basename(match.filePath)
      });
    }
    notify({ done: total, total, test, finished: true });
    // Test sends go to yourself; they don't belong in the sent reports.
    const logPath = test ? null : safeWriteSentLog(logEntries, report);
    const seconds = Math.round((Date.now() - started) / 1000);
    logSend(`Done: ${total} ${kind}${total === 1 ? '' : 's'} sent in ${seconds}s.${logPath ? ` Report: ${logPath}` : ''}`);
    notifyDone(`${total} ${kind}${total === 1 ? '' : 's'} sent`, test ? 'Check your inbox.' : 'A send report was saved.');
    return { success: true, logPath, sentIds };
  } catch (error) {
    // Still record whatever was sent before the failure.
    const logPath = test ? null : safeWriteSentLog(logEntries, report, error.message);
    const message = `Stopped after ${sentIds.length} of ${total} ${test ? 'test ' : ''}emails. ${error.message}`;
    logSend(`FAILED: ${message}`);
    notify({ done: sentIds.length, total, test, finished: true, failed: true });
    dialog.showErrorBox(test ? 'Test email error' : 'Email Error', test
      ? message
      : `${message}\n\nThe ${sentIds.length} email(s) that did go out are marked as sent and will not be sent again.`);
    return { success: false, message, logPath, sentIds };
  } finally {
    sendInProgress = false;
    if (powerSaveBlocker.isStarted(blockerId)) {
      powerSaveBlocker.stop(blockerId);
    }
  }
});

function formatDuration(ms) {
  const seconds = Math.max(0, Math.round(ms / 1000));
  if (seconds < 60) {
    return `${seconds}s`;
  }
  const minutes = Math.floor(seconds / 60);
  const rest = seconds % 60;
  return rest ? `${minutes}m ${rest}s` : `${minutes}m`;
}

ipcMain.handle('get-log-status', async () => {
  return { success: true, hasEntries: hasSentLogs() };
});

ipcMain.handle('open-sent-logs', async () => {
  if (!sentLogsDir) {
    sentLogsDir = path.join(app.getPath('userData'), 'sent-logs');
  }
  try {
    fs.mkdirSync(sentLogsDir, { recursive: true });
    await shell.openPath(sentLogsDir);
    return { success: true, path: sentLogsDir };
  } catch (error) {
    return { success: false, message: error.message };
  }
});

// ---------- Session copies of added files ----------
//
// The app never touches the files you add. Every file is copied into
// <userData>/upload-cache/ and only that copy is used for attaching. Copies are
// deleted when a file is removed from the list, on "Clear all", when the app
// quits, and (in case of a crash) again on the next start.

const MAX_IMPORT_FILES = 2000;

// macOS/Office leave hidden helper files in folders (.DS_Store, ~$file.docx).
function isIgnoredFileName(name) {
  return !name || name.startsWith('.') || name.startsWith('~$') || ['thumbs.db', 'desktop.ini'].includes(name.toLowerCase());
}

function isInsideUploadCache(targetPath) {
  if (!uploadCacheDir || !targetPath) {
    return false;
  }
  const relative = path.relative(uploadCacheDir, path.resolve(targetPath));
  return Boolean(relative) && !relative.startsWith('..') && !path.isAbsolute(relative);
}

function safeFileName(name) {
  return (path.basename(name || '') || 'file').replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').slice(-200) || 'file';
}

async function makeCopyDir() {
  const dir = path.join(uploadCacheDir, `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`);
  await fs.promises.mkdir(dir, { recursive: true });
  return dir;
}

// Expands files and folders (recursively) into a flat list of real file paths.
function expandPaths(targetPaths) {
  const results = [];
  const seen = new Set();
  function walk(entryPath) {
    const resolved = path.resolve(entryPath);
    if (seen.has(resolved) || isIgnoredFileName(path.basename(resolved)) || isInsideUploadCache(resolved)) {
      return;
    }
    seen.add(resolved);
    let stats;
    try {
      stats = fs.statSync(resolved);
    } catch (_) {
      return;
    }
    if (stats.isDirectory()) {
      fs.readdirSync(resolved).forEach((child) => walk(path.join(resolved, child)));
    } else if (stats.isFile()) {
      results.push(resolved);
      if (results.length > MAX_IMPORT_FILES) {
        throw new Error(`That is more than ${MAX_IMPORT_FILES} files. Drop the folder with just the feedback files.`);
      }
    }
  }
  targetPaths.filter(Boolean).forEach(walk);
  return results;
}

// Copies dropped files/folders into the session cache. Originals are only read.
ipcMain.handle('import-paths', async (_event, targetPaths = [], skipOriginals = []) => {
  try {
    if (!uploadCacheDir) {
      throw new Error('Upload cache not initialized.');
    }
    if (!Array.isArray(targetPaths)) {
      throw new Error('Invalid path list.');
    }
    const skip = new Set(Array.isArray(skipOriginals) ? skipOriginals : []);
    const found = expandPaths(targetPaths);
    const originals = found.filter((filePath) => !skip.has(filePath));
    const files = [];
    for (const originalPath of originals) {
      const dir = await makeCopyDir();
      const copyPath = path.join(dir, safeFileName(originalPath));
      // On APFS this makes an instant copy-on-write clone; elsewhere a normal copy.
      await fs.promises.copyFile(originalPath, copyPath, fs.constants.COPYFILE_FICLONE);
      files.push({ path: copyPath, name: path.basename(originalPath), originalPath });
    }
    files.sort((a, b) => a.name.localeCompare(b.name));
    return { success: true, files, skipped: found.length - originals.length };
  } catch (error) {
    return { success: false, message: error.message };
  }
});

// For files that have no location on disk (e.g. dragged out of an email).
ipcMain.handle('cache-uploaded-files', async (_event, files = []) => {
  try {
    if (!uploadCacheDir) {
      throw new Error('Upload cache not initialized.');
    }
    const cachedFiles = [];
    for (const file of files || []) {
      const data = file && file.data;
      const buffer = data ? Buffer.from(data instanceof ArrayBuffer ? new Uint8Array(data) : data) : null;
      if (!buffer || !buffer.length) {
        continue;
      }
      const dir = await makeCopyDir();
      const destination = path.join(dir, safeFileName(file.name));
      await fs.promises.writeFile(destination, buffer);
      cachedFiles.push({ path: destination, name: file.name || path.basename(destination), originalPath: '' });
    }
    return { success: true, files: cachedFiles };
  } catch (error) {
    return { success: false, message: error.message };
  }
});

// Deletes session copies. Refuses anything outside the upload cache, so an
// original file can never be deleted through this.
ipcMain.handle('discard-files', async (_event, copyPaths = []) => {
  let removed = 0;
  (Array.isArray(copyPaths) ? copyPaths : []).forEach((copyPath) => {
    if (!isInsideUploadCache(copyPath)) {
      return;
    }
    const dir = path.dirname(path.resolve(copyPath));
    try {
      fs.rmSync(isInsideUploadCache(dir) ? dir : copyPath, { recursive: true, force: true });
      removed += 1;
    } catch (error) {
      console.warn('Unable to remove cached copy', copyPath, error.message);
    }
  });
  return { success: true, removed };
});

function cleanupSessionFiles() {
  try {
    if (uploadCacheDir && fs.existsSync(uploadCacheDir)) {
      fs.rmSync(uploadCacheDir, { recursive: true, force: true });
    }
  } catch (_) {
    // ignore cleanup errors on quit
  }
  try {
    if (preparedAppleScriptPath) {
      fs.rmSync(path.dirname(preparedAppleScriptPath), { recursive: true, force: true });
    }
  } catch (_) {
    // ignore
  }
}

app.on('will-quit', cleanupSessionFiles);

// ---------- Templates ----------

// `keywords`: when most feedback file names contain one of these, the template
// is selected automatically (e.g. rubric_MMT1_CMD1B_Schols_Dean.pdf → MMT1).
function rubricTemplate(name, course, keywords) {
  return {
    name,
    subject: `Rubric ${course} - {{firstname}} {{lastname}}`,
    body: `Dag {{firstname}},\n\nHierbij ontvang je de rubric voor ${course}. Heb je hier vragen over, neem dan contact op met je docent.\n\nDit is een geautomatiseerd bericht.`,
    keywords
  };
}

const DEFAULT_TEMPLATES = [
  rubricTemplate('Rubric WS', 'Writing Skills', ['ws', 'writing skills']),
  rubricTemplate('Rubric MMT1', 'MMT1', ['mmt1', 'mmt', 'multimedia technology']),
  rubricTemplate('Rubric Visual Design', 'Visual Design', ['vd', 'visual design']),
  rubricTemplate('Rubric Photography', 'Photography', ['ph', 'photography', 'fotografie']),
  rubricTemplate('Rubric Design Research', 'Design Research', ['dr', 'design research'])
];

function readSettings() {
  try {
    return JSON.parse(fs.readFileSync(path.join(userDataDir, 'settings.json'), 'utf8')) || {};
  } catch {
    return {};
  }
}

function writeSettings(settings) {
  fs.writeFileSync(path.join(userDataDir, 'settings.json'), JSON.stringify(settings, null, 2), 'utf8');
}

// Adds each default template once. A default you delete stays deleted, and an
// existing template with the same name keeps its text (it only gains any
// default keywords it doesn't have yet).
function seedDefaultTemplates() {
  try {
    const settings = readSettings();
    const seeded = new Set(Array.isArray(settings.seededTemplates) ? settings.seededTemplates : []);
    const templates = readTemplates();
    let changed = false;
    DEFAULT_TEMPLATES.forEach((template) => {
      const existing = templates.find((entry) => entry.name === template.name);
      if (existing) {
        const current = Array.isArray(existing.keywords) ? existing.keywords : [];
        const missing = template.keywords.filter((keyword) => !current.includes(keyword));
        if (missing.length) {
          existing.keywords = current.concat(missing);
          changed = true;
        }
      } else if (!existing && !seeded.has(template.name)) {
        templates.push({ ...template });
        changed = true;
      }
      seeded.add(template.name);
    });
    if (changed) {
      writeTemplates(templates);
    }
    writeSettings({ ...settings, seededTemplates: Array.from(seeded) });
  } catch (error) {
    console.error('Unable to add default templates', error);
  }
}

ipcMain.handle('list-templates', async () => {
  const templates = readTemplates();
  return { success: true, templates };
});

ipcMain.handle('save-template', async (_event, template) => {
  try {
    if (!template || !template.name) {
      return { success: false, message: 'Template name is required.', templates: readTemplates() };
    }
    const current = readTemplates();
    const previous = current.find((entry) => entry.name === template.name);
    const templates = current.filter((entry) => entry.name !== template.name);
    templates.push({
      name: template.name,
      subject: template.subject || '',
      body: template.body || '',
      // Overwriting a template keeps its auto-select keywords.
      ...(previous && Array.isArray(previous.keywords) ? { keywords: previous.keywords } : {})
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
