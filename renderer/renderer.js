const excelInput = document.getElementById('excel-input');
const excelDropZone = document.getElementById('excel-drop-zone');
const dropZone = document.getElementById('drop-zone');
const feedbackInput = document.getElementById('feedback-input');
const fileList = document.getElementById('file-list');
const matchBody = document.getElementById('match-body');
const sendBtn = document.getElementById('send-btn');
const testBtn = document.getElementById('test-btn');
const openLogsBtn = document.getElementById('open-logs-btn');
const clearFilesBtn = document.getElementById('clear-files-btn');
const matchSummaryEl = document.getElementById('match-summary');
const unmatchedDetails = document.getElementById('unmatched-details');
const unmatchedSummary = document.getElementById('unmatched-summary');
const unmatchedList = document.getElementById('unmatched-list');
const attentionEl = document.getElementById('attention');
const attentionTitle = document.getElementById('attention-title');
const attentionList = document.getElementById('attention-list');
const placeholderWarningEl = document.getElementById('placeholder-warning');
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
const SKIP = '__skip__';
const KNOWN_PLACEHOLDERS = ['firstname', 'lastname', 'email', 'studentid'];
const EMAIL_PATTERN = /^[^\s@;,]+@[^\s@;,]+\.[^\s@;,]+$/;

let students = [];        // roster rows, each with a stable `key` and name `tokens`
let feedbackFiles = [];   // { path, name, cached }
let matches = [];         // student ↔ file pairs that can be sent
let unmatchedStudents = [];
let conflicts = [];       // ambiguous matches the user has to decide on
let fileStates = new Map(); // file path → { status, student? }
let rosterIssues = [];    // problems found in the roster itself
let blockedKeys = new Map(); // student key → reason this student can't be emailed
let messageTemplates = [];
let activeTemplateName = null;
let sending = false;
const selectionState = new Map();
// The user's answers to conflicts. Cleared when a new roster is loaded.
const fileChoices = new Map();    // file path → student key | SKIP
const studentChoices = new Map(); // student key → file path | SKIP
// Everything sent this session, so it isn't selected (and resent) again by accident.
const sentIds = new Set();

setRosterStatus('No roster loaded yet.');
setFeedbackStatus('No files loaded yet.');
updatePlaceholderWarning();

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
testBtn.addEventListener('click', handleTestSend);
openLogsBtn.addEventListener('click', handleOpenLogs);
clearFilesBtn.addEventListener('click', clearFeedbackFiles);
saveTemplateBtn.addEventListener('click', handleSaveTemplate);
// Dropping a file anywhere else would make the window navigate to it.
['dragover', 'drop'].forEach((type) => document.addEventListener(type, (event) => event.preventDefault()));
// Editing the message manually means it no longer equals the applied template.
[subjectInput, bodyInput].forEach((el) => el.addEventListener('input', () => {
  updatePlaceholderWarning();
  if (activeTemplateName) {
    activeTemplateName = null;
    renderTemplateButtons();
  }
}));
if (selectAllMatchesCheckbox) {
  selectAllMatchesCheckbox.addEventListener('change', () => {
    const checked = selectAllMatchesCheckbox.checked;
    // Already-sent and blocked rows are left alone; tick those one by one if you really mean it.
    getBulkSelectableMatches().forEach((match) => selectionState.set(match.id, checked));
    renderMatchTable();
    updateSelectionControls();
  });
}
if (window.electronAPI.onSendProgress) {
  window.electronAPI.onSendProgress(({ done, total, current, test, waitMs, remainingMs }) => {
    if (done < total && !progressFinished) {
      const left = remainingMs > 1500 ? ` · about ${formatDuration(remainingMs)} left` : '';
      const detail = waitMs
        ? `Next: ${current || ''} — short pause (max ~30 emails/minute)${left}`
        : `Now sending: ${current || ''}${left}`;
      updateProgress({ done, total, detail, test });
      setStatus(`Sending ${done + 1} of ${total}${current ? ` — ${current}` : ''}…`);
    }
  });
}
document.getElementById('send-progress-close').addEventListener('click', hideProgress);

// ---------- Progress panel ----------

let progressFinished = false;
let progressHideTimer = null;

function showProgress(total, test) {
  clearTimeout(progressHideTimer);
  progressFinished = false;
  const panel = document.getElementById('send-progress');
  panel.classList.remove('done', 'failed');
  document.getElementById('send-progress-close').hidden = true;
  panel.hidden = false;
  updateProgress({ done: 0, total, detail: 'Starting Outlook…', test });
}

function updateProgress({ done, total, detail, test }) {
  document.getElementById('send-progress-title').textContent = `Sending ${test ? 'test ' : ''}email${total === 1 ? '' : 's'}…`;
  document.getElementById('send-progress-count').textContent = `${done} of ${total} sent`;
  const bar = document.getElementById('send-progress-bar');
  bar.max = total || 1;
  bar.value = done;
  document.getElementById('send-progress-detail').textContent = detail || '';
}

function finishProgress({ sent, total, ok, message, test }) {
  progressFinished = true;
  const panel = document.getElementById('send-progress');
  panel.classList.add(ok ? 'done' : 'failed');
  document.getElementById('send-progress-title').textContent = ok
    ? `✓ ${sent} ${test ? 'test ' : ''}email${sent === 1 ? '' : 's'} sent`
    : '⚠ Sending stopped';
  document.getElementById('send-progress-count').textContent = `${sent} of ${total} sent`;
  const bar = document.getElementById('send-progress-bar');
  bar.max = total || 1;
  bar.value = sent;
  const detail = document.getElementById('send-progress-detail');
  detail.textContent = message || '';
  detail.title = message || '';
  document.getElementById('send-progress-close').hidden = false;
  // Success fades out by itself; an error stays until you close it.
  if (ok) {
    progressHideTimer = setTimeout(hideProgress, 8000);
  }
}

function formatDuration(ms) {
  const seconds = Math.max(0, Math.round(ms / 1000));
  if (seconds < 60) {
    return `${seconds} s`;
  }
  const minutes = Math.floor(seconds / 60);
  const rest = seconds % 60;
  return rest ? `${minutes} min ${rest} s` : `${minutes} min`;
}

function hideProgress() {
  clearTimeout(progressHideTimer);
  document.getElementById('send-progress').hidden = true;
}
showAppVersion();

async function showAppVersion() {
  try {
    const { version, electron } = await window.electronAPI.getAppInfo();
    const badge = document.getElementById('app-version');
    badge.textContent = `v${version}`;
    badge.title = `Version ${version} · Electron ${electron}`;
    badge.hidden = false;
    document.title = `Student Batch Mailer ${version}`;
  } catch (error) {
    console.warn('Unable to read app version', error);
  }
}

loadTemplates()
  .catch((error) => {
    console.error('Initial template load failed:', error);
  })
  .finally(() => {
    // Start with Writing Skills (or the first template); files can switch it later.
    const startTemplate = messageTemplates.find((template) => template.name === 'Rubric WS') || messageTemplates[0];
    if (startTemplate) {
      applyTemplate(startTemplate, { quiet: true });
    }
    updateMatches();
  });

// ---------- Roster ----------

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
    setRosterStatus('Please drop a .xlsx file.', 'error');
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
    const result = await window.electronAPI.parseExcel(loadEvent.target.result);
    if (!result.success) {
      setRosterStatus(result.message || 'Failed to read Excel file.', 'error');
      return;
    }
    students = result.students.map((student, index) => ({
      ...student,
      firstname: (student.firstname || '').trim(),
      lastname: (student.lastname || '').trim(),
      email: (student.email || '').trim(),
      studentid: (student.studentid || '').trim(),
      key: `s${index}`,
      tokens: getStudentNameTokens(student.firstname, student.lastname),
      tokenGroups: getStudentTokenGroups(student.firstname, student.lastname)
    }));
    // Old answers point at rows of the previous roster.
    fileChoices.clear();
    studentChoices.clear();
    analyzeRoster();
    const issueCount = rosterIssues.length;
    setRosterStatus(
      `Loaded ${students.length} students${issueCount ? ` — ${issueCount} problem${issueCount === 1 ? '' : 's'} in the roster, see step 4` : ''}.`,
      issueCount ? 'warning' : 'success'
    );
    const hasStudentIds = students.some((student) => student.studentid);
    document.getElementById('studentid-hint').hidden = !hasStudentIds;
    renderRosterPreview();
    updateMatches();
  };
  reader.onerror = () => setRosterStatus('Error reading file.', 'error');
  reader.readAsArrayBuffer(file);
}

// Finds roster problems that would make matching or sending go wrong.
function analyzeRoster() {
  rosterIssues = [];
  blockedKeys = new Map();
  const rowList = (list) => {
    const rows = list.map((s) => s.rowNumber);
    return rows.length > 1 ? `Rows ${rows.slice(0, -1).join(', ')} and ${rows[rows.length - 1]}` : `Row ${rows[0]}`;
  };
  const block = (student, reason) => {
    if (!blockedKeys.has(student.key)) {
      blockedKeys.set(student.key, reason);
    }
  };

  students.forEach((student) => {
    const label = studentName(student) || '(no name)';
    if (!student.firstname || !student.lastname || !student.tokens.length) {
      const missing = !student.firstname && !student.lastname ? 'first and last name are' : !student.firstname ? 'first name is' : 'last name is';
      rosterIssues.push({ text: `Row ${student.rowNumber} (${label}): ${missing} empty, so no file can ever be matched to this student.` });
    }
    if (!student.email) {
      rosterIssues.push({ text: `Row ${student.rowNumber} (${label}): no email address. This student will not be emailed.` });
      block(student, 'no email address');
    } else if (!EMAIL_PATTERN.test(student.email)) {
      rosterIssues.push({ text: `Row ${student.rowNumber} (${label}): "${student.email}" is not a valid email address. This student will not be emailed.` });
      block(student, 'invalid email address');
    }
  });

  const groupBy = (keyFn) => {
    const groups = new Map();
    students.forEach((student) => {
      const key = keyFn(student);
      if (!key) {
        return;
      }
      if (!groups.has(key)) {
        groups.set(key, []);
      }
      groups.get(key).push(student);
    });
    return Array.from(groups.values()).filter((group) => group.length > 1);
  };

  groupBy((s) => (EMAIL_PATTERN.test(s.email) ? s.email.toLowerCase() : '')).forEach((group) => {
    rosterIssues.push({ text: `${rowList(group)} share the email address ${group[0].email} (${group.map(studentName).join(' / ')}). Fix the roster; these students will not be emailed.` });
    group.forEach((student) => block(student, `email shared with another row`));
  });
  groupBy((s) => (s.tokens.length ? [...s.tokens].sort().join(' ') : '')).forEach((group) => {
    rosterIssues.push({ text: `${rowList(group)} have the same name (${studentName(group[0])}). Files for them can't be told apart, so you'll be asked to choose.` });
  });
  groupBy((s) => s.studentid.toLowerCase()).forEach((group) => {
    rosterIssues.push({ text: `${rowList(group)} have the same student ID (${group[0].studentid}).` });
  });
}

// ---------- Feedback files ----------

async function handleDrop(event) {
  event.preventDefault();
  dropZone.classList.remove('dragover');
  await addDroppedItems(event.dataTransfer);
}

async function handleFeedbackInputChange(event) {
  const files = Array.from(event.target.files || []);
  event.target.value = '';
  const paths = files.map((file) => window.electronAPI.getPathForFile(file));
  if (paths.length && paths.every(Boolean)) {
    await importPaths(paths);
  } else {
    await importPathlessFiles(files);
  }
}

function isIgnoredFileName(name) {
  return !name || name.startsWith('.') || name.startsWith('~$') || ['thumbs.db', 'desktop.ini'].includes(name.toLowerCase());
}

// Everything added is copied into the app's session cache first; only the copies
// are used. The originals are never modified, moved or deleted.
async function addDroppedItems(dataTransfer) {
  if (!dataTransfer) {
    return;
  }
  // Read everything synchronously: dataTransfer is emptied after the first await.
  const droppedFiles = Array.from(dataTransfer.files || []);
  const droppedPaths = droppedFiles.map((file) => window.electronAPI.getPathForFile(file));
  const entries = Array.from(dataTransfer.items || [])
    .map((item) => (item.webkitGetAsEntry ? item.webkitGetAsEntry() : null))
    .filter(Boolean);
  const uriPaths = extractFilePaths(dataTransfer);

  const paths = droppedFiles.length && droppedPaths.every(Boolean) ? droppedPaths : uriPaths;
  if (paths.length) {
    await importPaths(paths);
    return;
  }
  // Items without a location on disk (e.g. attachments dragged out of an email).
  let files = droppedFiles;
  if (entries.length) {
    const nested = await Promise.all(entries.map(entryToFiles));
    const flattened = nested.flat().filter(Boolean);
    if (flattened.length) {
      files = flattened;
    }
  }
  await importPathlessFiles(files);
}

async function importPaths(paths) {
  setFeedbackStatus('Copying files…', 'info');
  const known = feedbackFiles.map((file) => file.originalPath).filter(Boolean);
  let response;
  try {
    response = await window.electronAPI.importPaths(paths, known);
  } catch (error) {
    response = { success: false, message: error.message };
  }
  if (!response.success) {
    setFeedbackStatus(response.message || 'Unable to add the files.', 'error');
    return;
  }
  if (!response.files.length) {
    if (response.skipped) {
      setFeedbackStatus('Those files were already added.', 'info');
    } else {
      setFeedbackStatus('No usable files found.', 'error');
    }
    return;
  }
  await addFeedbackFiles(response.files);
}

async function importPathlessFiles(files) {
  const usable = (files || []).filter((file) => file && !isIgnoredFileName(file.name) && typeof file.arrayBuffer === 'function');
  if (!usable.length) {
    setFeedbackStatus('No usable files found.', 'error');
    return;
  }
  setFeedbackStatus('Copying files…', 'info');
  const cached = await cacheFilesWithoutPaths(usable);
  if (!cached.length) {
    setFeedbackStatus('Unable to access the selected files.', 'error');
    return;
  }
  await addFeedbackFiles(cached);
}

// Takes { path (session copy), name, originalPath } records.
async function addFeedbackFiles(newFiles) {
  const before = feedbackFiles.length;
  const { kept, dropped } = dedupeFiles(feedbackFiles.concat(newFiles.filter((file) => !isIgnoredFileName(file.name))));
  feedbackFiles = kept;
  discardCopies(dropped);
  updateMatches();
  if (feedbackFiles.length > before) {
    autoSelectTemplate();
  }
}

function discardCopies(files) {
  const copyPaths = files.map((file) => file.path).filter(Boolean);
  if (copyPaths.length) {
    window.electronAPI.discardFiles(copyPaths).catch((error) => console.warn('Unable to discard copies', error));
  }
}

function removeFeedbackFile(filePath) {
  discardCopies(feedbackFiles.filter((file) => file.path === filePath));
  feedbackFiles = feedbackFiles.filter((file) => file.path !== filePath);
  updateMatches();
}

function clearFeedbackFiles() {
  if (!feedbackFiles.length) {
    return;
  }
  if (feedbackFiles.length > 3 && !window.confirm(`Remove all ${feedbackFiles.length} files? (Only the app's copies are deleted; your original files stay where they are.)`)) {
    return;
  }
  discardCopies(feedbackFiles);
  feedbackFiles = [];
  updateMatches();
}

// The same original added twice is kept once (the newest copy). Two different
// files that share a name (from different folders) are both kept, and then show
// up as a choice instead of one silently disappearing. Files without an original
// location are compared by name.
function dedupeFiles(fileArray) {
  const seen = new Map();
  const dropped = [];
  fileArray.forEach((file) => {
    const key = file.originalPath ? `path:${file.originalPath}` : `name:${(file.name || '').toLowerCase()}`;
    if (seen.has(key)) {
      dropped.push(seen.get(key));
    }
    seen.set(key, file);
  });
  const kept = Array.from(seen.values()).sort((a, b) => a.name.localeCompare(b.name) || a.path.localeCompare(b.path));
  return { kept, dropped };
}

function fileDisplayPath(file) {
  return file.originalPath || `${file.name} (copied from a drag without a file location)`;
}

function updateFeedbackStatus() {
  const total = feedbackFiles.length;
  clearFilesBtn.hidden = total === 0;
  if (!total) {
    setFeedbackStatus('No files loaded yet.');
    return;
  }
  const label = total === 1 ? 'file' : 'files';
  if (!students.length) {
    setFeedbackStatus(`${total} ${label} loaded — load a roster to see matches.`, 'info');
    return;
  }
  const counts = { matched: 0, conflict: 0, replaced: 0, other: 0 };
  feedbackFiles.forEach((file) => {
    const status = (fileStates.get(file.path) || {}).status;
    if (status === 'matched') {
      counts.matched += 1;
    } else if (status === 'replaced') {
      counts.replaced += 1;
    } else if (status === 'conflict') {
      counts.conflict += 1;
    } else {
      counts.other += 1;
    }
  });
  const wordNote = counts.replaced ? `, ${counts.replaced} Word file${counts.replaced === 1 ? '' : 's'} skipped (PDF used)` : '';
  if (counts.matched + counts.replaced === total) {
    setFeedbackStatus(`${total} ${label} loaded, all matched${wordNote}.`, 'success');
    return;
  }
  const parts = [`${total} ${label} loaded`, `${counts.matched} matched`];
  if (counts.replaced) {
    parts.push(`${counts.replaced} Word file${counts.replaced === 1 ? '' : 's'} skipped (PDF used)`);
  }
  if (counts.conflict) {
    parts.push(`${counts.conflict} need${counts.conflict === 1 ? 's' : ''} your choice`);
  }
  if (counts.other) {
    parts.push(`${counts.other} not used`);
  }
  setFeedbackStatus(`${parts.join(', ')}.`, 'warning');
}

function renderFileList() {
  fileList.innerHTML = '';
  feedbackFiles.forEach((file) => {
    const state = fileStates.get(file.path) || { status: students.length ? 'unmatched' : 'pending' };
    const li = document.createElement('li');
    const name = document.createElement('span');
    name.className = 'file-name';
    name.textContent = file.name;
    name.title = fileDisplayPath(file);

    const badge = document.createElement('span');
    badge.className = 'badge';
    let itemClass = state.status;
    switch (state.status) {
      case 'matched':
        badge.textContent = `→ ${studentName(state.student)}${state.manual ? ' (your choice)' : ''}`;
        break;
      case 'conflict':
        badge.textContent = '⚠ needs your choice';
        break;
      case 'skipped':
        badge.textContent = 'skipped by you';
        itemClass = 'muted';
        break;
      case 'replaced':
        badge.textContent = `Word file — PDF sent to ${studentName(state.student)}`;
        itemClass = 'muted';
        break;
      case 'not-chosen':
        badge.textContent = `not used (other file chosen for ${studentName(state.student)})`;
        itemClass = 'muted';
        break;
      case 'pending':
        badge.textContent = 'waiting for roster';
        break;
      default:
        badge.textContent = 'no match';
        itemClass = 'unmatched';
    }
    badge.title = badge.textContent;
    li.className = `file-item ${itemClass}`;

    const removeBtn = document.createElement('button');
    removeBtn.type = 'button';
    removeBtn.className = 'icon-btn';
    removeBtn.textContent = '×';
    removeBtn.title = `Remove ${file.name}`;
    removeBtn.setAttribute('aria-label', `Remove ${file.name}`);
    removeBtn.addEventListener('click', () => removeFeedbackFile(file.path));

    li.append(name, badge, removeBtn);
    fileList.appendChild(li);
  });
}

// ---------- Matching ----------

// Matching rules:
//  1. A file is a candidate for a student when every word of the student's first
//     and last name appears in the file name.
//  2. A file that is a candidate for more than one student is NOT assigned until
//     the user chooses (conflict type "file").
//  3. A student that ends up with more than one file gets none of them until the
//     user chooses (conflict type "student").
//  The app never picks "the first one".
function computeMatches() {
  const result = { matches: [], unmatched: [], conflicts: [], fileStates: new Map() };
  if (!students.length) {
    return result;
  }
  const byKey = new Map(students.map((student) => [student.key, student]));
  const filesByStudent = new Map();
  const studentsInConflict = new Set();

  feedbackFiles.forEach((file) => {
    const fileTokens = tokenizeFileName(file.name);
    const candidates = students.filter((student) =>
      student.tokenGroups.length && student.tokenGroups.every((variants) => variants.some((token) => fileTokens.includes(token)))
    );
    const choice = fileChoices.get(file.path);
    const choiceValid = choice === SKIP || candidates.some((student) => student.key === choice);

    if (candidates.length > 1) {
      result.conflicts.push({
        type: 'file',
        file,
        candidates,
        suggestion: mostSpecificCandidate(candidates),
        choice: choiceValid ? choice : '',
        resolved: choiceValid
      });
      if (!choiceValid) {
        result.fileStates.set(file.path, { status: 'conflict' });
        candidates.forEach((student) => studentsInConflict.add(student.key));
        return;
      }
    }
    if (choiceValid && choice === SKIP) {
      result.fileStates.set(file.path, { status: 'skipped' });
      return;
    }
    const owner = choiceValid ? byKey.get(choice) : candidates.length === 1 ? candidates[0] : null;
    if (!owner) {
      result.fileStates.set(file.path, { status: 'unmatched' });
      return;
    }
    if (!filesByStudent.has(owner.key)) {
      filesByStudent.set(owner.key, []);
    }
    filesByStudent.get(owner.key).push({ file, manual: choiceValid && candidates.length > 1 });
  });

  students.forEach((student) => {
    let linked = filesByStudent.get(student.key) || [];
    if (!linked.length) {
      if (!studentsInConflict.has(student.key)) {
        result.unmatched.push(student);
      }
      return;
    }
    // Word and PDF version of the same rubric: always send the PDF.
    if (linked.some((entry) => isPdfFile(entry.file.name))) {
      linked.filter((entry) => isWordFile(entry.file.name))
        .forEach((entry) => result.fileStates.set(entry.file.path, { status: 'replaced', student }));
      linked = linked.filter((entry) => !isWordFile(entry.file.name));
    }
    // Only a Word file: never sent unless you explicitly say so.
    const wordOnly = linked.length === 1 && isWordFile(linked[0].file.name);
    let chosen = null;
    let manual = false;
    if (linked.length === 1 && !wordOnly) {
      chosen = linked[0];
      manual = chosen.manual;
    } else {
      const choice = studentChoices.get(student.key);
      const picked = choice && choice !== SKIP ? linked.find((entry) => entry.file.path === choice) : null;
      const resolved = choice === SKIP || Boolean(picked);
      result.conflicts.push({
        type: 'student',
        wordOnly,
        student,
        files: linked.map((entry) => entry.file),
        choice: resolved ? choice : '',
        resolved
      });
      if (!resolved) {
        linked.forEach((entry) => result.fileStates.set(entry.file.path, { status: 'conflict' }));
        return;
      }
      linked.forEach((entry) => {
        if (entry !== picked) {
          result.fileStates.set(entry.file.path, choice === SKIP ? { status: 'skipped' } : { status: 'not-chosen', student });
        }
      });
      if (!picked) {
        return; // the user chose not to send to this student
      }
      chosen = picked;
      manual = true;
    }
    result.fileStates.set(chosen.file.path, { status: 'matched', student, manual });
    result.matches.push({
      id: buildMatchId(student, chosen.file.path),
      key: student.key,
      rowNumber: student.rowNumber,
      firstname: student.firstname,
      lastname: student.lastname,
      email: student.email,
      studentid: student.studentid,
      fileName: chosen.file.name,
      filePath: chosen.file.path,
      originalPath: fileDisplayPath(chosen.file),
      manual,
      blockedReason: blockedKeys.get(student.key) || ''
    });
  });

  const byName = (a, b) => (a.lastname || '').localeCompare(b.lastname || '', 'nl')
    || (a.firstname || '').localeCompare(b.firstname || '', 'nl')
    || (a.rowNumber || 0) - (b.rowNumber || 0);
  result.matches.sort(byName);
  result.unmatched.sort(byName);
  return result;
}

// If one candidate's name contains all the others (e.g. "Anna de Vries-Bakker"
// vs "Anna de Vries" for anna-de-vries-bakker.pdf) it's a likely answer. It is only
// shown as a hint, never applied automatically.
function mostSpecificCandidate(candidates) {
  const sorted = [...candidates].sort((a, b) => b.tokens.length - a.tokens.length);
  const [best, ...rest] = sorted;
  const containsAll = rest.every((other) =>
    other.tokens.length < best.tokens.length && other.tokens.every((token) => best.tokens.includes(token))
  );
  return containsAll ? best.key : null;
}

function updateMatches() {
  const result = computeMatches();
  matches = result.matches;
  unmatchedStudents = result.unmatched;
  conflicts = result.conflicts;
  fileStates = result.fileStates;

  const currentIds = new Set(matches.map((match) => match.id));
  Array.from(selectionState.keys()).forEach((id) => {
    if (!currentIds.has(id)) {
      selectionState.delete(id);
    }
  });
  matches.forEach((match) => {
    if (match.blockedReason) {
      selectionState.set(match.id, false);
    } else if (!selectionState.has(match.id)) {
      selectionState.set(match.id, !sentIds.has(match.id));
    }
  });
  renderFileList();
  updateFeedbackStatus();
  renderAttention();
  renderMatchTable();
  renderUnmatchedStudents();
  updateSelectionControls();
}

function describeStudent(student) {
  const details = [`row ${student.rowNumber}`];
  if (student.email) {
    details.push(student.email);
  }
  if (student.studentid) {
    details.push(`ID ${student.studentid}`);
  }
  return `${studentName(student) || '(no name)'} (${details.join(', ')})`;
}

function describeConflict(conflict) {
  if (conflict.type === 'file') {
    return `${conflict.file.name} fits ${conflict.candidates.length} students: ${conflict.candidates.map(describeStudent).join('; ')}`;
  }
  if (conflict.wordOnly) {
    return `${describeStudent(conflict.student)} only has a Word file: ${conflict.files[0].name}`;
  }
  return `${describeStudent(conflict.student)} fits ${conflict.files.length} files: ${conflict.files.map((file) => file.name).join('; ')}`;
}

function getUnresolvedConflicts() {
  return conflicts.filter((conflict) => !conflict.resolved);
}

function renderAttention() {
  attentionList.innerHTML = '';
  const unresolved = getUnresolvedConflicts();
  const resolved = conflicts.filter((conflict) => conflict.resolved);
  if (!unresolved.length && !resolved.length && !rosterIssues.length) {
    attentionEl.hidden = true;
    return;
  }
  attentionEl.hidden = false;
  const openCount = unresolved.length + rosterIssues.length;
  attentionEl.classList.toggle('all-resolved', openCount === 0);
  attentionTitle.textContent = openCount
    ? `⚠ ${openCount} thing${openCount === 1 ? '' : 's'} need${openCount === 1 ? 's' : ''} your attention`
    : `✓ All matching questions answered (${resolved.length})`;

  unresolved.forEach((conflict) => attentionList.appendChild(renderConflictItem(conflict)));
  rosterIssues.forEach((issue) => {
    const li = document.createElement('li');
    li.className = 'attention-item roster';
    li.textContent = issue.text;
    attentionList.appendChild(li);
  });
  resolved.forEach((conflict) => attentionList.appendChild(renderConflictItem(conflict)));
}

function renderConflictItem(conflict) {
  const li = document.createElement('li');
  li.className = `attention-item${conflict.resolved ? ' resolved' : ' open'}`;
  const text = document.createElement('div');
  text.className = 'attention-text';
  const strong = document.createElement('strong');
  const select = document.createElement('select');
  const addOption = (value, label) => {
    const option = document.createElement('option');
    option.value = value;
    option.textContent = label;
    select.appendChild(option);
  };

  if (conflict.type === 'file') {
    strong.textContent = conflict.file.name;
    text.append(strong, ` fits ${conflict.candidates.length} students. Who should get this file?`);
    strong.title = conflict.file.path;
    addOption('', conflict.resolved ? '— undo my choice —' : '— choose a student —');
    conflict.candidates.forEach((student) => {
      addOption(student.key, `${describeStudent(student)}${conflict.suggestion === student.key ? ' — closest name match' : ''}`);
    });
    addOption(SKIP, 'Nobody — don\'t send this file');
    select.addEventListener('change', () => {
      if (select.value) {
        fileChoices.set(conflict.file.path, select.value);
      } else {
        fileChoices.delete(conflict.file.path);
      }
      updateMatches();
    });
  } else if (conflict.wordOnly) {
    strong.textContent = describeStudent(conflict.student);
    text.append(strong, ` only has a Word file (${conflict.files[0].name}), no PDF. Word files are not sent unless you say so.`);
    addOption('', conflict.resolved ? '— undo my choice —' : '— choose —');
    addOption(conflict.files[0].path, `Send the Word file anyway`);
    addOption(SKIP, 'Don\'t email this student (add a PDF first)');
    select.addEventListener('change', () => {
      if (select.value) {
        studentChoices.set(conflict.student.key, select.value);
      } else {
        studentChoices.delete(conflict.student.key);
      }
      updateMatches();
    });
  } else {
    strong.textContent = describeStudent(conflict.student);
    text.append(strong, ` fits ${conflict.files.length} files. Which one should be sent?`);
    addOption('', conflict.resolved ? '— undo my choice —' : '— choose a file —');
    conflict.files.forEach((file) => addOption(file.path, file.name));
    addOption(SKIP, 'None — don\'t email this student');
    select.addEventListener('change', () => {
      if (select.value) {
        studentChoices.set(conflict.student.key, select.value);
      } else {
        studentChoices.delete(conflict.student.key);
      }
      updateMatches();
    });
  }
  select.value = conflict.choice || '';
  li.append(text, select);
  return li;
}

function renderUnmatchedStudents() {
  if (!students.length || !feedbackFiles.length || !unmatchedStudents.length) {
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
    li.textContent = studentName(student) || student.email || `row ${student.rowNumber}`;
    unmatchedList.appendChild(li);
  });
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
    } else if (getUnresolvedConflicts().length) {
      cell.textContent = 'No confirmed matches yet. Answer the questions above.';
    } else {
      cell.textContent = 'None of the files match a student. Check the file names against the roster.';
    }
    row.appendChild(cell);
    matchBody.appendChild(row);
    return;
  }
  matches.forEach((match) => {
    const isSelected = Boolean(selectionState.get(match.id));
    const blocked = Boolean(match.blockedReason);
    const row = document.createElement('tr');
    row.className = `match-row${isSelected ? ' selected' : ''}${blocked ? ' blocked' : ''}`;

    const nameCell = document.createElement('td');
    nameCell.textContent = studentName(match);
    const addChip = (label, kind, title) => {
      const chip = document.createElement('span');
      chip.className = `chip ${kind}`;
      chip.textContent = label;
      if (title) {
        chip.title = title;
      }
      nameCell.append(' ', chip);
    };
    if (blocked) {
      addChip(match.blockedReason, 'danger', 'Fix this in the roster and load it again.');
    }
    if (match.manual) {
      addChip('your choice', 'info', 'You picked this match under “Needs your attention”.');
    }
    if (sentIds.has(match.id)) {
      addChip('✓ sent', 'success', 'Already sent in this session. Tick it again only if you want to resend.');
    }
    const emailCell = document.createElement('td');
    emailCell.textContent = match.email || '—';
    const fileCell = document.createElement('td');
    fileCell.className = 'file-cell';
    fileCell.textContent = match.fileName;
    fileCell.title = match.originalPath;

    const selectCell = document.createElement('td');
    selectCell.className = 'match-select';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = isSelected;
    checkbox.disabled = blocked;
    checkbox.setAttribute('aria-label', `Send to ${studentName(match)}`);
    selectCell.appendChild(checkbox);
    row.append(nameCell, emailCell, fileCell, selectCell);

    const setSelected = (next) => {
      if (blocked) {
        return;
      }
      selectionState.set(match.id, next);
      checkbox.checked = next;
      row.classList.toggle('selected', next);
      updateSelectionControls();
    };
    // The checkbox toggles natively; stop the click from bubbling so the row
    // handler doesn't immediately toggle it back.
    checkbox.addEventListener('click', (event) => event.stopPropagation());
    checkbox.addEventListener('change', () => setSelected(checkbox.checked));
    row.addEventListener('click', () => setSelected(!selectionState.get(match.id)));
    matchBody.appendChild(row);
  });
}

// ---------- Sending ----------

function getPlaceholderProblems(text) {
  const problems = [];
  const unknown = new Set();
  const withoutValid = (text || '').replace(/{{\s*([^{}]*?)\s*}}/g, (_all, name) => {
    if (!KNOWN_PLACEHOLDERS.includes(name.toLowerCase())) {
      unknown.add(`{{${name}}}`);
    }
    return '';
  });
  if (unknown.size) {
    problems.push(`Unknown placeholder${unknown.size === 1 ? '' : 's'} ${Array.from(unknown).join(', ')} will be sent literally. Use ${KNOWN_PLACEHOLDERS.map((name) => `{{${name}}}`).join(', ')}.`);
  }
  if (/[{}]/.test(withoutValid)) {
    problems.push('There is a stray { or } — check that every placeholder has two braces on both sides, like {{firstname}}.');
  }
  return problems;
}

function updatePlaceholderWarning() {
  const problems = getPlaceholderProblems(`${subjectInput.value}\n${bodyInput.value}`);
  if (!placeholderWarningEl) {
    return;
  }
  placeholderWarningEl.hidden = !problems.length;
  placeholderWarningEl.textContent = problems.length ? `⚠ ${problems.join(' ')}` : '';
}

function buildConfirmWarnings(selected) {
  const warnings = [];
  const unresolved = getUnresolvedConflicts();
  if (unresolved.length) {
    warnings.push({
      text: `${unresolved.length} matching question${unresolved.length === 1 ? ' is' : 's are'} still open. The students and files involved will NOT be emailed:`,
      details: unresolved.map(describeConflict)
    });
  }
  const resent = selected.filter((match) => sentIds.has(match.id));
  if (resent.length) {
    warnings.push({
      text: `${resent.length} of these ${resent.length === 1 ? 'was' : 'were'} already sent earlier in this session and will be sent AGAIN:`,
      details: resent.map((match) => `${studentName(match)} — ${match.fileName}`)
    });
  }
  const manual = selected.filter((match) => match.manual);
  if (manual.length) {
    warnings.push({
      text: `${manual.length} match${manual.length === 1 ? ' was' : 'es were'} chosen by you. Double-check:`,
      details: manual.map((match) => `${studentName(match)} ← ${match.fileName}`)
    });
  }
  const detected = detectTemplates();
  if (detected.mixed) {
    warnings.push({ text: `The file names mention more than one course (${detected.hits.map((hit) => `${hit.template.name}: ${hit.count}`).join(', ')}), but every student gets the same message.` });
  } else if (detected.best && detected.best.template.name !== activeTemplateName) {
    warnings.push({ text: `The files look like "${detected.best.template.name}", but the message ${activeTemplateName ? `uses template "${activeTemplateName}"` : 'is not that template'}.` });
  }
  const placeholderProblems = getPlaceholderProblems(`${subjectInput.value}\n${bodyInput.value}`);
  if (placeholderProblems.length) {
    warnings.push({ text: placeholderProblems.join(' ') });
  }
  if (!subjectInput.value.trim()) {
    warnings.push({ text: 'The subject is empty; the default subject will be used.' });
  }
  return warnings;
}

// Shows a large in-app confirmation listing every recipient. Resolves true/false.
function confirmSend(selected) {
  const dialog = document.getElementById('confirm-dialog');
  const body = document.getElementById('confirm-body');
  const title = document.getElementById('confirm-title');
  const subtitle = document.getElementById('confirm-subtitle');
  const warningsEl = document.getElementById('confirm-warnings');
  const okBtn = document.getElementById('confirm-ok');
  const cancelBtn = document.getElementById('confirm-cancel');

  const count = selected.length;
  title.textContent = `Send ${count} email${count === 1 ? '' : 's'} via Outlook?`;
  subtitle.textContent = `Subject: ${fillTemplate(subjectInput.value || DEFAULT_SUBJECT, selected[0])}`;
  okBtn.textContent = `Send ${count} email${count === 1 ? '' : 's'}`;
  if (count > 1) {
    subtitle.textContent += ` · takes about ${formatDuration((count - 1) * 2500)} (one email every 2.5 s). You can keep working in other apps; don't close this window or put the Mac to sleep.`;
  }

  const warnings = buildConfirmWarnings(selected);
  warningsEl.innerHTML = '';
  warningsEl.hidden = !warnings.length;
  warnings.forEach((warning) => {
    const block = document.createElement('div');
    block.className = 'confirm-warning';
    const p = document.createElement('p');
    p.textContent = `⚠ ${warning.text}`;
    block.appendChild(p);
    if (warning.details && warning.details.length) {
      const ul = document.createElement('ul');
      warning.details.forEach((detail) => {
        const li = document.createElement('li');
        li.textContent = detail;
        ul.appendChild(li);
      });
      block.appendChild(ul);
    }
    warningsEl.appendChild(block);
  });

  body.innerHTML = '';
  selected.forEach((match, index) => {
    const row = document.createElement('tr');
    [String(index + 1), studentName(match), match.email, match.fileName].forEach((value, column) => {
      const td = document.createElement('td');
      td.textContent = value;
      if (column === 0) {
        td.className = 'muted';
      }
      if (column === 3) {
        td.className = 'file-cell';
        td.title = match.originalPath;
      }
      row.appendChild(td);
    });
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
    // With warnings, put focus on Cancel so Enter doesn't send by reflex.
    (warnings.length ? cancelBtn : okBtn).focus();
  });
}

const DEFAULT_SUBJECT = 'Feedback for {{firstname}} {{lastname}}';
const DEFAULT_BODY = 'Please find your feedback attached.';

function fillTemplate(template, student) {
  const values = {
    firstname: student.firstname || '',
    lastname: student.lastname || '',
    email: student.email || '',
    studentid: student.studentid || ''
  };
  // A replacer function, so names containing "$" are inserted literally.
  return template.replace(/{{\s*(firstname|lastname|email|studentid)\s*}}/gi, (_all, name) => values[name.toLowerCase()]);
}

async function handleSend() {
  if (sending) {
    return;
  }
  const selectedMatches = getSelectedMatches();
  if (!selectedMatches.length) {
    setStatus('No matches selected to send.', 'error');
    return;
  }
  const confirmed = await confirmSend(selectedMatches);
  if (!confirmed) {
    return;
  }
  const plural = selectedMatches.length === 1 ? '' : 's';
  sending = true;
  updateSelectionControls();
  setStatus(`Sending ${selectedMatches.length} email${plural}…`);
  showProgress(selectedMatches.length, false);

  const subjectTemplate = subjectInput.value.trim() ? subjectInput.value : DEFAULT_SUBJECT;
  const bodyTemplate = bodyInput.value.trim() ? bodyInput.value : DEFAULT_BODY;
  const payloadMatches = selectedMatches.map((match) => ({
    id: match.id,
    firstname: match.firstname,
    lastname: match.lastname,
    email: match.email,
    studentid: match.studentid,
    fileName: match.fileName,
    filePath: match.filePath,
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
  let response;
  try {
    response = await window.electronAPI.sendEmails({
      matches: payloadMatches,
      report: {
        subjectTemplate,
        skipped: matches.filter((match) => !selectionState.get(match.id)).map(pick),
        unmatchedStudents: unmatchedStudents.map(pick),
        unmatchedFiles: feedbackFiles
          .filter((file) => !['matched', 'replaced'].includes((fileStates.get(file.path) || {}).status))
          .map((file) => file.name),
        problems: getUnresolvedConflicts().map(describeConflict).concat(rosterIssues.map((issue) => issue.text))
      }
    });
  } catch (error) {
    response = { success: false, message: error.message, sentIds: [] };
  } finally {
    sending = false;
  }

  // Mark whatever actually went out — also after a failure halfway — so it isn't resent.
  (response.sentIds || []).forEach((id) => {
    sentIds.add(id);
    selectionState.set(id, false);
  });
  if (response.success) {
    const logNote = response.logPath ? ` Log saved as ${response.logPath.split('/').pop()}.` : '';
    setStatus(`${selectedMatches.length} email${plural} sent.${logNote}`, 'success');
  } else {
    setStatus(response.message || 'Failed to send emails.', 'error');
  }
  finishProgress({
    sent: (response.sentIds || []).length,
    total: selectedMatches.length,
    ok: response.success,
    message: response.success ? 'A send report was saved (Open sent logs folder).' : response.message,
    test: false
  });
  renderMatchTable();
  updateSelectionControls();
}

// ---------- Test send ----------

const TEST_COUNT = 5;
const DEFAULT_TEST_ADDRESS = 'vincent.sijben@zuyd.nl';
const TEST_ADDRESS_KEY = 'sbm.testAddress';

// The first rows of the review list (as shown), skipping rows that can't be sent.
function getTestMatches() {
  return matches.filter((match) => !match.blockedReason).slice(0, TEST_COUNT);
}

function loadTestAddress() {
  try {
    return localStorage.getItem(TEST_ADDRESS_KEY) || DEFAULT_TEST_ADDRESS;
  } catch (_) {
    return DEFAULT_TEST_ADDRESS;
  }
}

function saveTestAddress(address) {
  try {
    localStorage.setItem(TEST_ADDRESS_KEY, address);
  } catch (_) {
    // not important
  }
}

// Asks where the test should go. Resolves to the address, or null on cancel.
function askTestAddress(picks) {
  const dialog = document.getElementById('test-dialog');
  const form = document.getElementById('test-form');
  const input = document.getElementById('test-address');
  const error = document.getElementById('test-error');
  const list = document.getElementById('test-list');
  const okBtn = document.getElementById('test-ok');
  const cancelBtn = document.getElementById('test-cancel');

  document.getElementById('test-title').textContent = `Send ${picks.length} test email${picks.length === 1 ? '' : 's'} to yourself`;
  okBtn.textContent = `Send ${picks.length} test email${picks.length === 1 ? '' : 's'}`;
  input.value = loadTestAddress();
  error.hidden = true;
  list.innerHTML = '';
  picks.forEach((match) => {
    const li = document.createElement('li');
    li.append(`${studentName(match)} `);
    const detail = document.createElement('span');
    detail.className = 'muted';
    detail.textContent = `— ${match.fileName}`;
    li.appendChild(detail);
    list.appendChild(li);
  });

  return new Promise((resolve) => {
    const cleanup = (result) => {
      form.removeEventListener('submit', onSubmit);
      cancelBtn.removeEventListener('click', onCancel);
      dialog.removeEventListener('close', onClose);
      if (dialog.open) {
        dialog.close();
      }
      resolve(result);
    };
    const onSubmit = (event) => {
      event.preventDefault();
      const address = input.value.trim();
      if (!EMAIL_PATTERN.test(address)) {
        error.textContent = 'Enter a valid email address.';
        error.hidden = false;
        input.focus();
        return;
      }
      cleanup(address);
    };
    const onCancel = () => cleanup(null);
    const onClose = () => cleanup(null); // Esc key
    form.addEventListener('submit', onSubmit);
    cancelBtn.addEventListener('click', onCancel);
    dialog.addEventListener('close', onClose);
    dialog.showModal();
    input.focus();
    input.select();
  });
}

async function handleTestSend() {
  if (sending) {
    return;
  }
  const picks = getTestMatches();
  if (!picks.length) {
    setStatus('There are no matches to test with yet.', 'error');
    return;
  }
  const address = await askTestAddress(picks);
  if (!address) {
    return;
  }
  saveTestAddress(address);
  const plural = picks.length === 1 ? '' : 's';
  sending = true;
  updateSelectionControls();
  setStatus(`Sending ${picks.length} test email${plural} to ${address}…`);
  showProgress(picks.length, true);

  const subjectTemplate = subjectInput.value.trim() ? subjectInput.value : DEFAULT_SUBJECT;
  const bodyTemplate = bodyInput.value.trim() ? bodyInput.value : DEFAULT_BODY;
  const payloadMatches = picks.map((match) => ({
    id: match.id,
    firstname: match.firstname,
    lastname: match.lastname,
    studentid: match.studentid,
    fileName: match.fileName,
    filePath: match.filePath,
    // Only the recipient changes; subject, text and attachment are the real ones.
    email: address,
    recipientName: `TEST – ${studentName(match)}`,
    subject: `[TEST → ${studentName(match)}] ${fillTemplate(subjectTemplate, match)}`,
    body: normalizeBodyText(`*** TEST — in the real send this email goes to ${studentName(match)} <${match.email}> ***\n\n${fillTemplate(bodyTemplate, match)}`)
  }));

  let response;
  try {
    response = await window.electronAPI.sendEmails({ matches: payloadMatches, test: true });
  } catch (error) {
    response = { success: false, message: error.message };
  } finally {
    sending = false;
  }
  // Test emails don't count as sent: the students got nothing.
  if (response.success) {
    setStatus(`${picks.length} test email${plural} sent to ${address}. Check your inbox; students received nothing.`, 'success');
  } else {
    setStatus(`Test failed: ${response.message || 'unknown error'}`, 'error');
  }
  finishProgress({
    sent: (response.sentIds || []).length,
    total: picks.length,
    ok: response.success,
    message: response.success ? `Sent to ${address} — check your inbox.` : response.message,
    test: true
  });
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
  return matches.filter((match) => !match.blockedReason && selectionState.get(match.id));
}

function getBulkSelectableMatches() {
  return matches.filter((match) => !match.blockedReason && !sentIds.has(match.id));
}

function updateSelectionControls() {
  const selectedCount = getSelectedMatches().length;
  const hasMatches = matches.length > 0;
  sendBtn.disabled = sending || !hasMatches || selectedCount === 0;
  const testCount = getTestMatches().length;
  testBtn.disabled = sending || testCount === 0;
  testBtn.textContent = testCount ? `Send test (${testCount}) to myself` : 'Send test to myself';
  sendBtn.textContent = sending
    ? 'Sending…'
    : selectedCount
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
      const open = getUnresolvedConflicts().length;
      if (open) {
        parts.push(`${open} to decide`);
      }
      matchSummaryEl.textContent = parts.join(' · ');
    }
  }

  if (selectAllMatchesCheckbox) {
    const bulk = getBulkSelectableMatches();
    const bulkSelected = bulk.filter((match) => selectionState.get(match.id)).length;
    selectAllMatchesCheckbox.disabled = sending || !bulk.length;
    selectAllMatchesCheckbox.checked = bulk.length > 0 && bulkSelected === bulk.length;
    selectAllMatchesCheckbox.indeterminate = bulkSelected > 0 && bulkSelected < bulk.length;
  }
}

function setTemplateStatus(message, type) {
  if (!templateStatusEl) {
    return;
  }
  templateStatusEl.textContent = message || '';
  templateStatusEl.className = `template-status${type ? ` ${type}` : ''}`;
}

// ---------- Name helpers ----------

const WORD_EXTENSIONS = ['doc', 'docx', 'docm', 'dot', 'dotx', 'odt', 'rtf', 'pages'];

function fileExtension(name) {
  const match = /\.([^.]+)$/.exec(name || '');
  return match ? match[1].toLowerCase() : '';
}

function isWordFile(name) {
  return WORD_EXTENSIONS.includes(fileExtension(name));
}

function isPdfFile(name) {
  return fileExtension(name) === 'pdf';
}

// One entry per name word, listing the spellings that count as a match:
// "Schütt" → ['schutt', 'schuett'], so schutt.pdf and schuett.pdf both work.
function getStudentTokenGroups(first, last) {
  if (!getStudentNameTokens(first, last).length) {
    return [];
  }
  const lower = `${first || ''} ${last || ''}`.toLowerCase().normalize('NFC');
  const plain = splitTokens(foldDiacritics(lower));
  const german = splitTokens(foldDiacritics(lower.replace(/ä/g, 'ae').replace(/ö/g, 'oe').replace(/ü/g, 'ue')));
  const groups = new Map();
  plain.forEach((token, index) => {
    if (!groups.has(token)) {
      groups.set(token, new Set([token]));
    }
    if (german.length === plain.length) {
      groups.get(token).add(german[index]);
    }
  });
  return Array.from(groups.values()).map((set) => Array.from(set));
}

function splitTokens(value) {
  return value.replace(/[^a-z0-9]+/g, ' ').split(' ').filter(Boolean);
}

function studentName(student) {
  return student ? `${student.firstname || ''} ${student.lastname || ''}`.trim() : '';
}

function getStudentNameTokens(first, last) {
  const firstTokens = tokenizeNameValue(first);
  const lastTokens = tokenizeNameValue(last);
  if (!firstTokens.length || !lastTokens.length) {
    return [];
  }
  return Array.from(new Set([...firstTokens, ...lastTokens]));
}

function tokenizeFileName(name) {
  const withoutExtension = (name || '').toString().replace(/\.[^.]+$/, '');
  return tokenizeNameValue(withoutExtension);
}

function tokenizeNameValue(value) {
  return tokenizeLowerString(foldDiacritics((value || '').toString().toLowerCase()));
}

// "José Müller-Ørsted" → "jose muller-orsted". macOS stores file names in
// decomposed form (é = e + ◌́), so both sides are normalised the same way.
function foldDiacritics(value) {
  return value
    .normalize('NFKD')
    .replace(/\p{M}/gu, '')
    .replace(/ß/g, 'ss')
    .replace(/æ/g, 'ae')
    .replace(/œ/g, 'oe')
    .replace(/[øǿ]/g, 'o')
    .replace(/ł/g, 'l')
    .replace(/đ/g, 'd')
    .replace(/ı/g, 'i');
}

function tokenizeLowerString(value) {
  const parts = value
    .replace(/[^a-z0-9]+/g, ' ')
    .split(' ')
    .filter(Boolean);
  return Array.from(new Set(parts));
}

// Based on email + file, so a re-loaded roster still recognises what was sent.
function buildMatchId(student, filePath) {
  return [(student.email || student.key).toLowerCase(), filePath].join('|');
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
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
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
    const keywords = templateKeywords(template);
    applyBtn.title = `${template.name === activeTemplateName ? 'Template in use' : `Use template "${template.name}"`}${keywords.length ? `\nSelected automatically when file names contain: ${keywords.join(', ')}` : ''}`;
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

function applyTemplate(template, { quiet = false } = {}) {
  subjectInput.value = template.subject || '';
  bodyInput.value = template.body || '';
  activeTemplateName = template.name;
  renderTemplateButtons();
  updatePlaceholderWarning();
  if (!quiet) {
    setTemplateStatus(`Using template "${template.name}".`, 'success');
  }
}

// ---------- Template detection from file names ----------

const GENERIC_TEMPLATE_WORDS = ['rubric', 'rubrics', 'feedback', 'template', 'beoordeling'];

// Words that, when found in file names, point at this template. Defaults carry
// explicit keywords; for your own templates the name is used ("Rubric MMT2" → mmt2).
function templateKeywords(template) {
  if (Array.isArray(template.keywords) && template.keywords.length) {
    return template.keywords;
  }
  const words = tokenizeNameValue(template.name).filter((word) => !GENERIC_TEMPLATE_WORDS.includes(word));
  return words.length ? [words.join(' ')] : [];
}

function fileNameHasKeyword(fileName, keyword) {
  const sequence = splitTokens(foldDiacritics(fileName.toLowerCase().replace(/\.[^.]+$/, '')));
  const words = tokenizeNameValue(keyword);
  if (!words.length) {
    return false;
  }
  // "visual design" matches visual_design.pdf and VisualDesign.pdf.
  if (sequence.includes(words.join(''))) {
    return true;
  }
  for (let i = 0; i + words.length <= sequence.length; i += 1) {
    if (words.every((word, offset) => sequence[i + offset] === word)) {
      return true;
    }
  }
  return false;
}

// Which templates do the file names point at? Word files that were replaced by
// a PDF don't count twice.
function detectTemplates() {
  const files = feedbackFiles.filter((file) => (fileStates.get(file.path) || {}).status !== 'replaced');
  const hits = messageTemplates
    .map((template) => {
      const keywords = templateKeywords(template);
      const count = files.filter((file) => keywords.some((keyword) => fileNameHasKeyword(file.name, keyword))).length;
      return { template, count };
    })
    .filter((hit) => hit.count > 0)
    .sort((a, b) => b.count - a.count);
  const total = files.length;
  // A course only counts when it shows up in at least 20% of the files, so one
  // stray "vd" (van der) or "dr" in a file name doesn't look like a second course.
  const threshold = Math.max(1, Math.ceil(total * 0.2));
  const significant = hits.filter((hit) => hit.count >= threshold);
  const best = significant.length === 1 && significant[0].count >= total / 2 ? significant[0] : null;
  return { hits: significant, total, best, mixed: significant.length > 1 };
}

function messageIsUntouched() {
  if (activeTemplateName) {
    return true;
  }
  const subject = subjectInput.value.trim();
  const body = bodyInput.value.trim();
  return (!subject && !body) || messageTemplates.some((t) => t.subject === subjectInput.value && t.body === bodyInput.value);
}

function autoSelectTemplate() {
  const { hits, total, best, mixed } = detectTemplates();
  if (mixed) {
    setTemplateStatus(`⚠ The file names mention more than one course: ${hits.map((hit) => `${hit.template.name} (${hit.count} files)`).join(', ')}. The template was not changed — check it before sending.`, 'error');
    return;
  }
  if (!best) {
    return;
  }
  const name = best.template.name;
  const reason = `${best.count} of ${total} file names contain “${templateKeywords(best.template).join('” / “')}”`;
  if (activeTemplateName === name) {
    setTemplateStatus(`Template "${name}" matches the files (${reason}).`, 'success');
  } else if (messageIsUntouched()) {
    applyTemplate(best.template, { quiet: true });
    setTemplateStatus(`Template "${name}" selected automatically: ${reason}.`, 'success');
  } else {
    setTemplateStatus(`⚠ The files look like "${name}" (${reason}), but you edited the message, so it was not replaced. Click "${name}" to use it.`, 'error');
  }
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

// Plain text with \n line breaks; Outlook gets it as-is (no HTML).
function normalizeBodyText(text) {
  if (!text) {
    return '';
  }
  return text.replace(/\r\n?/g, '\n');
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
