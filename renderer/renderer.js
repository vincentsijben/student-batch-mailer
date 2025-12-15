const excelInput = document.getElementById('excel-input');
const excelDropZone = document.getElementById('excel-drop-zone');
const dropZone = document.getElementById('drop-zone');
const feedbackInput = document.getElementById('feedback-input');
const fileList = document.getElementById('file-list');
const matchBody = document.getElementById('match-body');
const sendBtn = document.getElementById('send-btn');
const exportLogBtn = document.getElementById('export-log-btn');
const openDataBtn = document.getElementById('open-data-btn');
const clearLogBtn = document.getElementById('clear-log-btn');
const templateNameInput = document.getElementById('template-name');
const saveTemplateBtn = document.getElementById('save-template-btn');
const templateButtonsContainer = document.getElementById('template-buttons');
const rosterStatusEl = document.getElementById('roster-status');
const feedbackStatusEl = document.getElementById('feedback-status');
const selectAllMatchesCheckbox = document.getElementById('select-all-matches');
const statusEl = document.getElementById('status');
const templateStatusEl = document.getElementById('template-status');
const subjectInput = document.getElementById('subject');
const bodyInput = document.getElementById('body');
const rosterPreview = document.getElementById('roster-preview');
if (rosterPreview) {
  rosterPreview.style.display = 'none';
}

let students = [];
let feedbackFiles = [];
let matches = [];
let messageTemplates = [];
const selectionState = new Map();

subjectInput.value = 'Rubric Writing Skills - {{firstname}} {{lastname}}';
bodyInput.value = 'Dag {{firstname}},\n\nHierbij ontvang je de rubric voor Writing Skills. Heb je hier vragen over, neem dan contact op met je docent.\n\nDit is een geautomatiseerd bericht.';
setRosterStatus('No roster loaded yet.');
setFeedbackStatus('No files loaded yet.');


excelInput.addEventListener('change', handleExcelUpload);
excelDropZone.addEventListener('click', () => excelInput.click());
excelDropZone.addEventListener('dragover', (event) => {
  event.preventDefault();
  excelDropZone.classList.add('dragover');
});
excelDropZone.addEventListener('dragleave', () => excelDropZone.classList.remove('dragover'));
excelDropZone.addEventListener('drop', handleExcelDrop);
dropZone.addEventListener('dragover', (event) => {
  event.preventDefault();
  dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
dropZone.addEventListener('drop', handleDrop);
dropZone.addEventListener('click', () => feedbackInput.click());
feedbackInput.addEventListener('change', handleFeedbackInputChange);
sendBtn.addEventListener('click', handleSend);
exportLogBtn.addEventListener('click', handleExportLog);
openDataBtn.addEventListener('click', handleOpenData);
clearLogBtn.addEventListener('click', handleClearLog);
saveTemplateBtn.addEventListener('click', handleSaveTemplate);
if (selectAllMatchesCheckbox) {
  selectAllMatchesCheckbox.addEventListener('change', () => {
    const checked = selectAllMatchesCheckbox.checked;
    matches.forEach((match) => {
      selectionState.set(match.id, checked);
    });
    renderMatchTable();
    updateSelectionControls();
  });
}
refreshExportAvailability();
loadTemplates()
  .catch((error) => {
    console.error('Initial template load failed:', error);
  })
  .finally(() => {
    initializeDefaultTemplate();
    updateSelectionControls();
  });

function handleExcelUpload(event) {
  const file = event.target.files[0];
  event.target.value = '';
  processExcelFile(file);
}

function handleExcelDrop(event) {
  event.preventDefault();
  excelDropZone.classList.remove('dragover');
  const files = Array.from(event.dataTransfer.files || []);
  if (!files.length) {
    return;
  }
  const file = files.find((entry) => entry.name && entry.name.toLowerCase().endsWith('.xlsx'));
  if (!file) {
    setStatus('Please select a .xlsx file.', 'error');
    return;
  }
  processExcelFile(file);
}

function processExcelFile(file) {
  if (!file) {
    return;
  }
  if (!file.name.toLowerCase().endsWith('.xlsx')) {
    setRosterStatus('Please select a .xlsx file.', 'error');
    return;
  }
  const reader = new FileReader();
  reader.onload = async (loadEvent) => {
    setRosterStatus('Parsing roster…', 'info');
    const arrayBuffer = loadEvent.target.result;
    const result = await window.electronAPI.parseExcel(arrayBuffer);
    if (!result.success) {
      setRosterStatus(result.message || 'Failed to read Excel file.', 'error');
      return;
    }
    students = result.students;
    setRosterStatus(`Loaded ${students.length} students.`, 'success');
    renderRosterPreview();
    updateMatches();
  };
  reader.onerror = () => setRosterStatus('Error reading file.', 'error');
  reader.readAsArrayBuffer(file);
}

async function handleDrop(event) {
  event.preventDefault();
  dropZone.classList.remove('dragover');
  const files = await collectDroppedFiles(event.dataTransfer);
  await addFeedbackFiles(files);
}

async function handleFeedbackInputChange(event) {
  const files = Array.from(event.target.files || []);
  await addFeedbackFiles(files);
  event.target.value = '';
}

async function addFeedbackFiles(files) {
  if (!files || !files.length) {
    setFeedbackStatus('No files selected.', 'error');
    return;
  }
  const resolvedFiles = [];
  const pendingCache = [];
  files.forEach((file) => {
    if (file && file.path && file.name) {
      resolvedFiles.push({ path: file.path, name: file.name });
    } else if (file && file.name && typeof file.arrayBuffer === 'function') {
      pendingCache.push(file);
    }
  });
  if (pendingCache.length) {
    setFeedbackStatus('Preparing dropped files…', 'info');
    const cached = await cacheFilesWithoutPaths(pendingCache);
    resolvedFiles.push(...cached);
  }
  if (!resolvedFiles.length) {
    setFeedbackStatus('Unable to access selected files.', 'error');
    return;
  }
  feedbackFiles = dedupeFiles(feedbackFiles.concat(resolvedFiles));
  const fileCount = feedbackFiles.length;
  const label = fileCount === 1 ? 'file' : 'files';
  setFeedbackStatus(`Loaded ${fileCount} ${label}.`, 'success');
  renderFileList();
  updateMatches();
}

function dedupeFiles(fileArray) {
  const seen = new Map();
  fileArray.forEach((file) => {
    if (!seen.has(file.path)) {
      seen.set(file.path, file);
    }
  });
  return Array.from(seen.values()).sort((a, b) => a.name.localeCompare(b.name));
}

async function collectDroppedFiles(dataTransfer) {
  if (!dataTransfer) {
    return [];
  }
  const items = Array.from(dataTransfer.items || []);
  if (items.length) {
    const entryPromises = items.map(async (item) => {
      const entry = item.webkitGetAsEntry ? item.webkitGetAsEntry() : null;
      if (!entry) {
        const file = item.getAsFile ? item.getAsFile() : null;
        return file ? [file] : [];
      }
      return entryToFiles(entry);
    });
    const nestedEntries = await Promise.all(entryPromises);
    const flattened = nestedEntries.flat().filter(Boolean);
    if (flattened.length) {
      return flattened;
    }
  }
  const fileList = Array.from(dataTransfer.files || []);
  const usable = fileList.filter((file) => Boolean(file.path));
  const needsExpansion = !fileList.length || usable.length !== fileList.length;
  let expanded = [];
  if (needsExpansion) {
    const uriPaths = extractFilePaths(dataTransfer);
    if (uriPaths.length && window.electronAPI.expandPaths) {
      try {
        const response = await window.electronAPI.expandPaths(uriPaths);
        if (response.success && Array.isArray(response.files)) {
          expanded = response.files.map((file) => ({
            path: file.path,
            name: file.name
          }));
        } else if (response.message) {
          console.warn('Unable to expand dropped items:', response.message);
        }
      } catch (error) {
        console.error('Failed to expand dropped items:', error);
      }
    }
  }
  if (usable.length || expanded.length) {
    return usable.concat(expanded);
  }
  return fileList;
}

function renderFileList() {
  fileList.innerHTML = '';
  feedbackFiles.forEach((file) => {
    const li = document.createElement('li');
    li.textContent = file.name;
    fileList.appendChild(li);
  });
}

function updateMatches() {
  matches = computeMatches();
  const currentIds = new Set(matches.map((match) => match.id));
  Array.from(selectionState.keys()).forEach((id) => {
    if (!currentIds.has(id)) {
      selectionState.delete(id);
    }
  });
  matches.forEach((match) => {
    if (!selectionState.has(match.id)) {
      selectionState.set(match.id, true);
    }
  });
  renderMatchTable();
  updateSelectionControls();
}

function renderRosterPreview() {
  rosterPreview.innerHTML = '';
  if (!students.length) {
    rosterPreview.style.display = 'none';
    return;
  }
  rosterPreview.style.display = 'block';

  const firstLabel = students.some((student) => student.sourceLanguage === 'dutch') ? 'Voornaam' : 'Firstname';
  const lastLabel = students.some((student) => student.sourceLanguage === 'dutch') ? 'Achternaam' : 'Lastname';
  const table = document.createElement('table');
  table.innerHTML = `
    <thead>
      <tr>
        <th>${firstLabel}</th>
        <th>${lastLabel}</th>
        <th>Email</th>
      </tr>
    </thead>
  `;
  const tbody = document.createElement('tbody');
  students.forEach((student) => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td>${student.firstname || ''}</td>
      <td>${student.lastname || ''}</td>
      <td>${student.email || ''}</td>
    `;
    tbody.appendChild(row);
  });
  table.appendChild(tbody);
  rosterPreview.appendChild(table);
}

function computeMatches() {
  const used = new Set();
  const normalizedFiles = feedbackFiles.map((file, index) => ({
    ...file,
    index,
    tokens: tokenizeFileName(file.name || '')
  }));

  return students
    .map((student) => {
      const first = (student.firstname || '').trim();
      const last = (student.lastname || '').trim();
      const email = (student.email || '').trim();
      const studentId = (student.studentid || '').trim();
      const nameTokens = getStudentNameTokens(first, last);

      const match = nameTokens.length
        ? normalizedFiles.find((file) => {
            if (used.has(file.index)) {
              return false;
            }
            return nameTokens.every((token) => file.tokens.includes(token));
          })
        : null;

      if (match) {
        const matchId = buildMatchId(first, last, match.path);
        used.add(match.index);
        return {
          id: matchId,
          firstname: first,
          lastname: last,
          email,
          studentid: studentId,
          fileName: match.name,
          filePath: match.path
        };
      }
      return {
        firstname: first,
        lastname: last,
        email,
        studentid: studentId,
        fileName: ''
      };
    })
    .filter((record) => record.filePath)
    .sort((a, b) => {
      const lastCompare = (a.lastname || '').localeCompare(b.lastname || '');
      if (lastCompare !== 0) {
        return lastCompare;
      }
      return (a.firstname || '').localeCompare(b.firstname || '');
    });
}

function renderMatchTable() {
  matchBody.innerHTML = '';
  matches.forEach((match) => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td>${match.firstname} ${match.lastname}</td>
      <td>${match.email}</td>
      <td>${match.fileName}</td>
    `;
    const selectCell = document.createElement('td');
    selectCell.className = 'match-select';
    const label = document.createElement('label');
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = Boolean(selectionState.get(match.id));
    checkbox.addEventListener('change', () => {
      selectionState.set(match.id, checkbox.checked);
      renderMatchTable();
      updateSelectionControls();
    });
    label.appendChild(checkbox);
    selectCell.appendChild(label);
    row.appendChild(selectCell);
    matchBody.appendChild(row);
  });
}

function fillTemplate(template, student) {
  return template
    .replace(/{{\s*firstname\s*}}/gi, student.firstname || '')
    .replace(/{{\s*lastname\s*}}/gi, student.lastname || '')
    .replace(/{{\s*email\s*}}/gi, student.email || '')
    .replace(/{{\s*studentid\s*}}/gi, student.studentid || '');
}

async function handleSend() {
  if (!matches.length) {
    return;
  }
  const selectedMatches = getSelectedMatches();
  if (!selectedMatches.length) {
    setStatus('No matches selected to send.', 'error');
    return;
  }
  const plural = selectedMatches.length === 1 ? '' : 's';
  const confirmed = window.confirm(`Send ${selectedMatches.length} email${plural}?`);
  if (!confirmed) {
    return;
  }
  setStatus('Sending emails…');
  sendBtn.disabled = true;

  const subjectTemplate = subjectInput.value || 'Feedback for {{firstname}} {{lastname}}';
  const bodyTemplate = bodyInput.value || 'Please find your feedback attached.';
  const payloadMatches = selectedMatches.map((match) => ({
    ...match,
    subject: fillTemplate(subjectTemplate, match),
    body: normalizeBodyText(fillTemplate(bodyTemplate, match))
  }));

  const response = await window.electronAPI.sendEmails({ matches: payloadMatches });
  if (response.success) {
    setStatus('Emails sent successfully.', 'success');
    if (typeof response.hasLogEntries !== 'undefined') {
      setExportEnabled(response.hasLogEntries);
    }
  } else {
    setStatus(response.message || 'Failed to send emails.', 'error');
    if (typeof response.hasLogEntries !== 'undefined') {
      setExportEnabled(response.hasLogEntries);
    }
  }
  updateSelectionControls();
}

function setStatus(message, type) {
  statusEl.textContent = message || '';
  statusEl.className = type ? type : '';
}

function setRosterStatus(message, type) {
  if (!rosterStatusEl) {
    return;
  }
  rosterStatusEl.textContent = message || '';
  rosterStatusEl.className = `status-note${type ? ` ${type}` : ''}`;
  if (!students.length) {
    rosterPreview.style.display = 'none';
  }
}

function setFeedbackStatus(message, type) {
  if (!feedbackStatusEl) {
    return;
  }
  feedbackStatusEl.textContent = message || '';
  feedbackStatusEl.className = `status-note${type ? ` ${type}` : ''}`;
}

function getSelectedMatches() {
  return matches.filter((match) => selectionState.get(match.id));
}

function updateSelectionControls() {
  const selectedCount = getSelectedMatches().length;
  const hasMatches = matches.length > 0;
  sendBtn.disabled = !hasMatches || selectedCount === 0;
  if (selectAllMatchesCheckbox) {
    if (!hasMatches) {
      selectAllMatchesCheckbox.checked = false;
      selectAllMatchesCheckbox.indeterminate = false;
      selectAllMatchesCheckbox.disabled = true;
    } else {
      selectAllMatchesCheckbox.disabled = false;
      const allSelected = selectedCount === matches.length;
      selectAllMatchesCheckbox.checked = allSelected;
      selectAllMatchesCheckbox.indeterminate = selectedCount > 0 && selectedCount < matches.length;
    }
  }
}

function setTemplateStatus(message, type) {
  if (!templateStatusEl) {
    return;
  }
  templateStatusEl.textContent = message || '';
  templateStatusEl.className = `template-status${type ? ` ${type}` : ''}`;
}

async function handleExportLog() {
  setStatus('Exporting log…');
  const response = await window.electronAPI.exportLog();
  if (response.success) {
    setStatus(`Log exported to ${response.filePath}`, 'success');
  } else {
    setStatus(response.message || 'Export failed.', 'error');
  }
  if (typeof response.hasLogEntries !== 'undefined') {
    setExportEnabled(response.hasLogEntries);
  }
}

async function refreshExportAvailability() {
  try {
    const response = await window.electronAPI.getLogStatus();
    setExportEnabled(Boolean(response && response.hasEntries));
  } catch {
    setExportEnabled(false);
  }
}

function setExportEnabled(enabled) {
  exportLogBtn.disabled = !enabled;
  clearLogBtn.disabled = !enabled;
}

function getStudentNameTokens(first, last) {
  const firstTokens = tokenizeNameValue(first);
  const lastTokens = tokenizeNameValue(last);
  if (!firstTokens.length || !lastTokens.length) {
    return [];
  }
  const combined = new Set([...firstTokens, ...lastTokens]);
  return Array.from(combined);
}

function tokenizeFileName(name) {
  const lower = (name || '').toString().toLowerCase();
  const withoutExtension = lower.replace(/\.[^.]+$/, '');
  return tokenizeLowerString(withoutExtension);
}

function tokenizeNameValue(value) {
  const lower = (value || '').toString().toLowerCase();
  return tokenizeLowerString(lower);
}

function tokenizeLowerString(value) {
  const parts = value
    .replace(/[^a-z0-9]+/g, ' ')
    .split(' ')
    .map((part) => part.trim())
    .filter(Boolean);
  return Array.from(new Set(parts));
}

function buildMatchId(first, last, filePath) {
  return [
    (first || '').toLowerCase(),
    (last || '').toLowerCase(),
    (filePath || '').toLowerCase()
  ].join('|');
}

async function handleOpenData() {
  setStatus('Opening data folder…');
  const response = await window.electronAPI.openUserData();
  if (response.success) {
    setStatus(`Opened ${response.path}`, 'success');
  } else {
    setStatus(response.message || 'Failed to open folder.', 'error');
  }
}

async function handleClearLog() {
  const confirmed = window.confirm('This will remove all sent email records. Continue?');
  if (!confirmed) {
    return;
  }
  setStatus('Clearing sent log…');
  const response = await window.electronAPI.clearLog();
  if (response.success) {
    setStatus('Sent log cleared.', 'success');
    setExportEnabled(false);
  } else {
    setStatus(response.message || 'Failed to clear log.', 'error');
    if (typeof response.hasEntries !== 'undefined') {
      setExportEnabled(response.hasEntries);
    }
  }
}

async function loadTemplates() {
  try {
    const response = await window.electronAPI.listTemplates();
    if (response.success) {
      messageTemplates = response.templates || [];
      renderTemplateButtons();
      console.log(`Loaded ${messageTemplates.length} templates.`);
    } else {
      throw new Error(response.message || 'Failed to load templates.');
    }
  } catch (error) {
    console.error('Unable to load templates.', error);
    messageTemplates = [];
    renderTemplateButtons();
  }
}

async function initializeDefaultTemplate() {
  const defaultName = 'Rubric WS';
  const exists = messageTemplates.some((template) => template.name === defaultName);
  if (exists) {
    return;
  }
  const payload = {
    name: defaultName,
    subject: 'Rubric Writing Skills - {{firstname}} {{lastname}}',
    body: 'Dag {{firstname}},\n\nHierbij ontvang je de rubric voor Writing Skills. Heb je hier vragen over, neem dan contact op met je docent.\n\nDit is een geautomatiseerd bericht.'
  };
  try {
    const response = await window.electronAPI.saveTemplate(payload);
    if (response.success) {
      messageTemplates = response.templates || [];
      renderTemplateButtons();
      console.log('Default template "Rubric WS" created.');
    }
  } catch (error) {
    console.error('Failed to create default template.', error);
  }
}

function renderTemplateButtons() {
  templateButtonsContainer.innerHTML = '';
  if (!messageTemplates.length) {
    const placeholder = document.createElement('p');
    placeholder.className = 'placeholder';
    placeholder.textContent = 'No templates saved yet.';
    templateButtonsContainer.appendChild(placeholder);
    return;
  }
  messageTemplates.forEach((template) => {
    const wrapper = document.createElement('div');
    wrapper.className = 'template-entry';

    const applyBtn = document.createElement('button');
    applyBtn.type = 'button';
    applyBtn.className = 'apply';
    applyBtn.textContent = template.name;
    applyBtn.addEventListener('click', () => applyTemplate(template));

    const deleteBtn = document.createElement('button');
    deleteBtn.type = 'button';
    deleteBtn.className = 'delete';
    deleteBtn.textContent = '×';
    deleteBtn.setAttribute('aria-label', `Delete ${template.name}`);
    deleteBtn.addEventListener('click', () => deleteTemplate(template.name));

    wrapper.appendChild(applyBtn);
    wrapper.appendChild(deleteBtn);
    templateButtonsContainer.appendChild(wrapper);
  });
}

function applyTemplate(template) {
  subjectInput.value = template.subject || '';
  bodyInput.value = template.body || '';
  setTemplateStatus(`Loaded template "${template.name}".`, 'success');
}

async function handleSaveTemplate() {
  const name = (templateNameInput.value || '').trim();
  if (!name) {
    setStatus('Template name is required.', 'error');
    console.warn('Template save aborted: missing name.');
    return;
  }
  const payload = {
    name,
    subject: subjectInput.value || '',
    body: bodyInput.value || ''
  };
  setStatus('Saving template…');
  console.log('Saving template', payload);
  try {
    const response = await window.electronAPI.saveTemplate(payload);
    if (response.success) {
      templateNameInput.value = '';
      await loadTemplates();
      setStatus(`Template "${name}" saved.`, 'success');
      console.log(`Template "${name}" saved.`);
    } else {
      setStatus(response.message || 'Failed to save template.', 'error');
      console.error('Template save failed:', response);
    }
  } catch (error) {
    setStatus(error.message || 'Failed to save template.', 'error');
    console.error('Template save error:', error);
  }
}

async function deleteTemplate(name) {
  const confirmed = window.confirm(`Delete template "${name}"?`);
  if (!confirmed) {
    console.warn('Template delete canceled by user.');
    return;
  }
  setStatus('Removing template…');
  try {
    const response = await window.electronAPI.deleteTemplate(name);
    if (response.success) {
      await loadTemplates();
      setStatus(`Template "${name}" removed.`, 'success');
      console.log(`Template "${name}" removed.`);
    } else {
      setStatus(response.message || 'Failed to delete template.', 'error');
      console.error('Template delete failed:', response);
    }
  } catch (error) {
    setStatus(error.message || 'Failed to delete template.', 'error');
    console.error('Template delete error:', error);
  }
}

function normalizeBodyText(text) {
  if (!text) {
    return '';
  }
  let normalized = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  normalized = normalized.replace(/<br\s*\/?>/gi, '\n');
  return normalized.replace(/\n/g, '\r');
}

function extractFilePaths(dataTransfer) {
  if (!dataTransfer || typeof dataTransfer.getData !== 'function') {
    return [];
  }
  const raw = dataTransfer.getData('text/uri-list') || '';
  return raw
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line && !line.startsWith('#') && line.startsWith('file://'))
    .map((line) => {
      try {
        const url = new URL(line);
        return decodeURI(url.pathname);
      } catch {
        return null;
      }
    })
    .filter(Boolean);
}

async function cacheFilesWithoutPaths(fileList) {
  if (!window.electronAPI || typeof window.electronAPI.cacheUploadedFiles !== 'function') {
    return [];
  }
  try {
    const payload = await Promise.all(
      fileList.map(async (file) => ({
        name: file.name || 'file',
        data: await file.arrayBuffer()
      }))
    );
    const response = await window.electronAPI.cacheUploadedFiles(payload);
    if (response.success && Array.isArray(response.files)) {
      return response.files;
    }
    if (response.message) {
      console.error('Cache upload failed:', response.message);
    }
    return [];
  } catch (error) {
    console.error('Cache upload error:', error);
    return [];
  }
}

async function entryToFiles(entry) {
  if (entry.isFile) {
    const file = await getFileFromEntry(entry);
    return file ? [file] : [];
  }
  if (entry.isDirectory) {
    const reader = entry.createReader();
    const entries = await readAllDirectoryEntries(reader);
    const childFiles = await Promise.all(entries.map(entryToFiles));
    return childFiles.flat();
  }
  return [];
}

function readAllDirectoryEntries(reader) {
  return new Promise((resolve) => {
    const entries = [];
    function readBatch() {
      reader.readEntries((batch) => {
        if (!batch.length) {
          resolve(entries);
          return;
        }
        entries.push(...batch);
        readBatch();
      }, () => resolve(entries));
    }
    readBatch();
  });
}

function getFileFromEntry(entry) {
  return new Promise((resolve) => {
    entry.file(
      (file) => resolve(file),
      () => resolve(null)
    );
  });
}
