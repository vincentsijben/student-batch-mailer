const excelInput = document.getElementById('excel-input');
const excelDropZone = document.getElementById('excel-drop-zone');
const dropZone = document.getElementById('drop-zone');
const feedbackInput = document.getElementById('feedback-input');
const fileList = document.getElementById('file-list');
const matchBody = document.getElementById('match-body');
const sendBtn = document.getElementById('send-btn');
const openLogsBtn = document.getElementById('open-logs-btn');
const clearFilesBtn = document.getElementById('clear-files-btn');
const matchSummaryEl = document.getElementById('match-summary');
const unmatchedDetails = document.getElementById('unmatched-details');
const unmatchedSummary = document.getElementById('unmatched-summary');
const unmatchedList = document.getElementById('unmatched-list');
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

const ROSTER_PREVIEW_LIMIT = 8;

let students = [];
let feedbackFiles = [];
let matches = [];
let unmatchedStudents = [];
let messageTemplates = [];
let activeTemplateName = null;
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
openLogsBtn.addEventListener('click', handleOpenLogs);
clearFilesBtn.addEventListener('click', clearFeedbackFiles);
saveTemplateBtn.addEventListener('click', handleSaveTemplate);
// Editing the message manually means it no longer equals the applied template.
[subjectInput, bodyInput].forEach((el) => el.addEventListener('input', () => {
  if (activeTemplateName) {
    activeTemplateName = null;
    renderTemplateButtons();
  }
}));
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
loadTemplates()
  .catch((error) => {
    console.error('Initial template load failed:', error);
  })
  .finally(async () => {
    await initializeDefaultTemplate();
    // The fields are pre-filled with the default template, so mark it active.
    const defaultTemplate = messageTemplates.find((template) => template.name === 'Rubric WS');
    if (defaultTemplate && subjectInput.value === defaultTemplate.subject && bodyInput.value === defaultTemplate.body) {
      activeTemplateName = defaultTemplate.name;
      renderTemplateButtons();
    }
    updateMatches();
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
    const hasStudentIds = students.some((student) => (student.studentid || '').trim());
    document.getElementById('studentid-hint').hidden = !hasStudentIds;
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
  updateMatches();
}

function removeFeedbackFile(filePath) {
  feedbackFiles = feedbackFiles.filter((file) => file.path !== filePath);
  updateMatches();
}

function clearFeedbackFiles() {
  if (!feedbackFiles.length) {
    return;
  }
  if (feedbackFiles.length > 3 && !window.confirm(`Remove all ${feedbackFiles.length} files?`)) {
    return;
  }
  feedbackFiles = [];
  updateMatches();
}

function updateFeedbackStatus() {
  const total = feedbackFiles.length;
  clearFilesBtn.hidden = total === 0;
  if (!total) {
    setFeedbackStatus('No files loaded yet.');
    return;
  }
  const matchedPaths = new Set(matches.map((match) => match.filePath));
  const matchedCount = feedbackFiles.filter((file) => matchedPaths.has(file.path)).length;
  const unmatchedCount = total - matchedCount;
  const label = total === 1 ? 'file' : 'files';
  if (!students.length) {
    setFeedbackStatus(`${total} ${label} loaded — load a roster to see matches.`, 'info');
  } else if (unmatchedCount === 0) {
    setFeedbackStatus(`${total} ${label} loaded, all matched.`, 'success');
  } else {
    setFeedbackStatus(`${total} ${label} loaded, ${matchedCount} matched, ${unmatchedCount} without a student.`, unmatchedCount ? 'warning' : 'success');
  }
}

// Files dropped without a filesystem path are copied into a fresh cache folder
// every time, so the same file dropped twice gets two different paths. Dedupe
// on the file name instead and keep the most recently added copy.
function dedupeFiles(fileArray) {
  const seen = new Map();
  fileArray.forEach((file) => {
    const key = (file.name || file.path || '').toLowerCase();
    seen.set(key, file);
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
  const matchByPath = new Map(matches.map((match) => [match.filePath, match]));
  feedbackFiles.forEach((file) => {
    const match = matchByPath.get(file.path);
    const li = document.createElement('li');
    li.className = `file-item ${match ? 'matched' : students.length ? 'unmatched' : 'pending'}`;

    const name = document.createElement('span');
    name.className = 'file-name';
    name.textContent = file.name;
    name.title = file.path;

    const badge = document.createElement('span');
    badge.className = 'badge';
    if (match) {
      badge.textContent = `→ ${match.firstname} ${match.lastname}`.trim();
    } else if (students.length) {
      badge.textContent = 'no match';
    } else {
      badge.textContent = 'waiting for roster';
    }

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'icon-btn';
    removeBtn.textContent = '×';
    removeBtn.title = `Remove ${file.name}`;
    removeBtn.setAttribute('aria-label', `Remove ${file.name}`);
    removeBtn.addEventListener('click', () => removeFeedbackFile(file.path));

    li.appendChild(name);
    li.appendChild(badge);
    li.appendChild(removeBtn);
    fileList.appendChild(li);
  });
}

function updateMatches() {
  const result = computeMatches();
  matches = result.matches;
  unmatchedStudents = result.unmatched;
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
  renderFileList();
  updateFeedbackStatus();
  renderMatchTable();
  renderUnmatchedStudents();
  updateSelectionControls();
}

function renderUnmatchedStudents() {
  if (!students.length || !unmatchedStudents.length) {
    unmatchedDetails.hidden = true;
    unmatchedDetails.open = false;
    return;
  }
  unmatchedDetails.hidden = false;
  const count = unmatchedStudents.length;
  unmatchedSummary.textContent = `${count} student${count === 1 ? '' : 's'} without a matching file`;
  unmatchedList.innerHTML = '';
  unmatchedStudents.forEach((student) => {
    const li = document.createElement('li');
    li.textContent = `${student.firstname} ${student.lastname}`.trim() || student.email;
    unmatchedList.appendChild(li);
  });
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
  students.slice(0, ROSTER_PREVIEW_LIMIT).forEach((student) => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td title="${escapeHtml(student.firstname || '')}">${escapeHtml(student.firstname || '')}</td>
      <td title="${escapeHtml(student.lastname || '')}">${escapeHtml(student.lastname || '')}</td>
      <td title="${escapeHtml(student.email || '')}">${escapeHtml(student.email || '')}</td>
    `;
    tbody.appendChild(row);
  });
  table.appendChild(tbody);
  rosterPreview.appendChild(table);
  if (students.length > ROSTER_PREVIEW_LIMIT) {
    const more = document.createElement('p');
    more.className = 'placeholder';
    more.textContent = `…and ${students.length - ROSTER_PREVIEW_LIMIT} more`;
    rosterPreview.appendChild(more);
  }
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function computeMatches() {
  const used = new Set();
  const normalizedFiles = feedbackFiles.map((file, index) => ({
    ...file,
    index,
    tokens: tokenizeFileName(file.name || '')
  }));

  const records = students
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
    });

  const byName = (a, b) => {
    const lastCompare = (a.lastname || '').localeCompare(b.lastname || '');
    if (lastCompare !== 0) {
      return lastCompare;
    }
    return (a.firstname || '').localeCompare(b.firstname || '');
  };

  return {
    matches: records.filter((record) => record.filePath).sort(byName),
    unmatched: records.filter((record) => !record.filePath).sort(byName)
  };
}

function renderMatchTable() {
  matchBody.innerHTML = '';
  if (!matches.length) {
    const row = document.createElement('tr');
    row.className = 'empty-row';
    const cell = document.createElement('td');
    cell.colSpan = 4;
    if (!students.length && !feedbackFiles.length) {
      cell.textContent = 'Load a roster and add feedback files to see matches here.';
    } else if (!students.length) {
      cell.textContent = 'Load a roster (step 1) to match the files.';
    } else if (!feedbackFiles.length) {
      cell.textContent = 'Add feedback files (step 2) to match them to students.';
    } else {
      cell.textContent = 'None of the files match a student. Check the file names against the roster.';
    }
    row.appendChild(cell);
    matchBody.appendChild(row);
    return;
  }
  matches.forEach((match) => {
    const isSelected = Boolean(selectionState.get(match.id));
    const row = document.createElement('tr');
    row.className = `match-row${isSelected ? ' selected' : ''}`;
    row.innerHTML = `
      <td>${escapeHtml(`${match.firstname} ${match.lastname}`.trim())}</td>
      <td>${escapeHtml(match.email)}</td>
      <td class="file-cell" title="${escapeHtml(match.filePath)}">${escapeHtml(match.fileName)}</td>
    `;
    const selectCell = document.createElement('td');
    selectCell.className = 'match-select';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = isSelected;
    selectCell.appendChild(checkbox);
    row.appendChild(selectCell);

    const setSelected = (next) => {
      selectionState.set(match.id, next);
      checkbox.checked = next;
      row.classList.toggle('selected', next);
      updateSelectionControls();
    };
    // The checkbox toggles natively; stop the click from bubbling so the row
    // handler doesn't immediately toggle it back.
    checkbox.addEventListener('click', (event) => event.stopPropagation());
    checkbox.addEventListener('change', () => setSelected(checkbox.checked));
    // Clicking anywhere else on the row toggles too.
    row.addEventListener('click', () => setSelected(!selectionState.get(match.id)));
    matchBody.appendChild(row);
  });
}

// Shows a large in-app confirmation listing every recipient. Resolves true/false.
function confirmSend(selected) {
  const dialog = document.getElementById('confirm-dialog');
  const body = document.getElementById('confirm-body');
  const title = document.getElementById('confirm-title');
  const subtitle = document.getElementById('confirm-subtitle');
  const okBtn = document.getElementById('confirm-ok');
  const cancelBtn = document.getElementById('confirm-cancel');

  const count = selected.length;
  title.textContent = `Send ${count} email${count === 1 ? '' : 's'} via Outlook?`;
  subtitle.textContent = `Subject: ${fillTemplate(subjectInput.value || '', selected[0])}`;
  okBtn.textContent = `Send ${count} email${count === 1 ? '' : 's'}`;

  body.innerHTML = '';
  selected.forEach((match, index) => {
    const row = document.createElement('tr');
    row.innerHTML = `
      <td class="muted">${index + 1}</td>
      <td>${escapeHtml(`${match.firstname} ${match.lastname}`.trim())}</td>
      <td>${escapeHtml(match.email)}</td>
      <td class="file-cell" title="${escapeHtml(match.fileName)}">${escapeHtml(match.fileName)}</td>
    `;
    body.appendChild(row);
  });

  return new Promise((resolve) => {
    const cleanup = (result) => {
      okBtn.removeEventListener('click', onOk);
      cancelBtn.removeEventListener('click', onCancel);
      dialog.removeEventListener('close', onClose);
      dialog.removeEventListener('click', onBackdrop);
      if (dialog.open) {
        dialog.close();
      }
      resolve(result);
    };
    const onOk = () => cleanup(true);
    const onCancel = () => cleanup(false);
    const onClose = () => cleanup(false); // Esc key
    const onBackdrop = (event) => {
      if (event.target === dialog) {
        cleanup(false);
      }
    };
    okBtn.addEventListener('click', onOk);
    cancelBtn.addEventListener('click', onCancel);
    dialog.addEventListener('close', onClose);
    dialog.addEventListener('click', onBackdrop);
    dialog.showModal();
    okBtn.focus();
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
  const confirmed = await confirmSend(selectedMatches);
  if (!confirmed) {
    return;
  }
  setStatus(`Sending ${selectedMatches.length} email${plural}…`);
  sendBtn.disabled = true;

  const subjectTemplate = subjectInput.value || 'Feedback for {{firstname}} {{lastname}}';
  const bodyTemplate = bodyInput.value || 'Please find your feedback attached.';
  const payloadMatches = selectedMatches.map((match) => ({
    ...match,
    subject: fillTemplate(subjectTemplate, match),
    body: normalizeBodyText(fillTemplate(bodyTemplate, match))
  }));

  const pick = (record) => ({
    firstname: record.firstname,
    lastname: record.lastname,
    email: record.email,
    studentid: record.studentid,
    fileName: record.fileName || ''
  });
  const response = await window.electronAPI.sendEmails({
    matches: payloadMatches,
    report: {
      subjectTemplate,
      skipped: matches.filter((match) => !selectionState.get(match.id)).map(pick),
      unmatchedStudents: unmatchedStudents.map(pick),
      unmatchedFiles: feedbackFiles
        .filter((file) => !matches.some((match) => match.filePath === file.path))
        .map((file) => file.name)
    }
  });
  if (response.success) {
    const logNote = response.logPath ? ` Log saved as ${response.logPath.split('/').pop()}.` : '';
    setStatus(`${selectedMatches.length} email${plural} sent.${logNote}`, 'success');
    // Deselect what was just sent so a second click can't resend by accident.
    selectedMatches.forEach((match) => selectionState.set(match.id, false));
    renderMatchTable();
  } else {
    setStatus(response.message || 'Failed to send emails.', 'error');
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
  sendBtn.textContent = selectedCount
    ? `Send ${selectedCount} email${selectedCount === 1 ? '' : 's'}`
    : 'Send emails';

  if (matchSummaryEl) {
    if (!students.length) {
      matchSummaryEl.textContent = '';
    } else {
      const parts = [`${matches.length} of ${students.length} students matched`];
      if (hasMatches) {
        parts.push(`${selectedCount} selected`);
      }
      matchSummaryEl.textContent = parts.join(' · ');
    }
  }

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

async function handleOpenLogs() {
  setStatus('Opening sent logs folder…');
  const response = await window.electronAPI.openSentLogs();
  if (response.success) {
    setStatus(`Opened ${response.path}`, 'success');
  } else {
    setStatus(response.message || 'Failed to open folder.', 'error');
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
    applyBtn.className = `apply${template.name === activeTemplateName ? ' active' : ''}`;
    applyBtn.textContent = template.name;
    applyBtn.title = template.name === activeTemplateName ? 'Template in use' : `Use template "${template.name}"`;
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
  activeTemplateName = template.name;
  renderTemplateButtons();
  setTemplateStatus(`Using template "${template.name}".`, 'success');
}

async function handleSaveTemplate() {
  const name = (templateNameInput.value || '').trim();
  if (!name) {
    setTemplateStatus('Give the template a name first.', 'error');
    templateNameInput.focus();
    return;
  }
  const payload = {
    name,
    subject: subjectInput.value || '',
    body: bodyInput.value || ''
  };
  setTemplateStatus('Saving template…');
  try {
    const response = await window.electronAPI.saveTemplate(payload);
    if (response.success) {
      templateNameInput.value = '';
      activeTemplateName = name;
      await loadTemplates();
      setTemplateStatus(`Template "${name}" saved.`, 'success');
    } else {
      setTemplateStatus(response.message || 'Failed to save template.', 'error');
      console.error('Template save failed:', response);
    }
  } catch (error) {
    setTemplateStatus(error.message || 'Failed to save template.', 'error');
    console.error('Template save error:', error);
  }
}

async function deleteTemplate(name) {
  const confirmed = window.confirm(`Delete template "${name}"?`);
  if (!confirmed) {
    return;
  }
  setTemplateStatus('Removing template…');
  try {
    const response = await window.electronAPI.deleteTemplate(name);
    if (response.success) {
      if (activeTemplateName === name) {
        activeTemplateName = null;
      }
      await loadTemplates();
      setTemplateStatus(`Template "${name}" removed.`, 'success');
    } else {
      setTemplateStatus(response.message || 'Failed to delete template.', 'error');
      console.error('Template delete failed:', response);
    }
  } catch (error) {
    setTemplateStatus(error.message || 'Failed to delete template.', 'error');
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
