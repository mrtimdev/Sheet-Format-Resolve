'use strict';

const $ = (id) => document.getElementById(id);
const form = $('cleaner');
const dropzone = $('dropzone');
const input = $('file');
const mode = $('mode');
const headerMode = $('header-mode');
const headerRows = $('header-rows');
const previewSheet = $('preview-sheet');
const options = [...form.querySelectorAll('input[type="checkbox"]')];
const accepted = /\.(xlsx|xls)$/i;
let file = null;
let busy = false;
let downloadUrl = null;
let dragDepth = 0;
let preview = null;
let previewRequest = null;
let previewMessage = '';
let previewFailed = false;
let previewLocation = { sheetIndex: 0, rowOffset: 0, columnOffset: 0 };

function renderMode() {
  const columns = mode.value === 'columns';
  $('column-options').hidden = !columns;
  $('cell-options').hidden = columns;
  headerMode.disabled = busy || !columns;
  $('manual-header').hidden = headerMode.value === 'auto';
  headerRows.disabled = busy || !columns || headerMode.value === 'auto';
  previewSheet.disabled = busy;
  $('preview-retry').disabled = busy;
  $('mode-title').textContent = columns ? 'Keep every row' : 'Keep cell positions';
  $('mode-note').textContent = columns
    ? 'Columns with data stay in their original order.'
    : 'Matching values are cleared in place.';
}

mode.addEventListener('change', () => {
  error();
  clearResult();
  renderMode();
  renderPreview();
});
renderMode();

headerMode.addEventListener('change', () => {
  renderMode();
  clearResult();
  renderPreview();
});
headerRows.addEventListener('input', () => {
  clearResult();
  renderPreview();
});
previewSheet.addEventListener('change', () => loadPreview(Number(previewSheet.value)));
$('preview-retry').addEventListener('click', () => {
  const { sheetIndex, rowOffset, columnOffset } = previewLocation;
  loadPreview(sheetIndex, rowOffset, columnOffset);
});
const previewNavigation = [
  ['previous-rows', p => p.rowOffset > 0, p => [Math.max(0, p.rowOffset - p.pageRows), p.columnOffset]],
  ['next-rows', p => p.moreRows, p => [p.rowOffset + p.rows.length, p.columnOffset]],
  ['previous-columns', p => p.columnOffset > 0, p => [p.rowOffset, Math.max(0, p.columnOffset - p.pageColumns)]],
  ['next-columns', p => p.moreColumns, p => [p.rowOffset, p.columnOffset + p.columns.length]]
];
for (const [id, , destination] of previewNavigation) {
  $(id).addEventListener('click', () => loadPreview(preview.sheetIndex, ...destination(preview)));
}

function clearPreview() {
  previewRequest?.abort();
  previewRequest = null;
  preview = null;
  previewMessage = '';
  previewFailed = false;
  previewSheet.replaceChildren();
  $('preview-table').replaceChildren();
}

async function loadPreview(sheetIndex = 0, rowOffset = 0, columnOffset = 0) {
  if (!file || busy) return;
  previewRequest?.abort();
  const request = new AbortController();
  previewRequest = request;
  previewLocation = { sheetIndex, rowOffset, columnOffset };
  previewFailed = false;
  previewMessage = 'Reading worksheet…';
  renderPreview();
  const data = new FormData();
  data.set('file', file);
  data.set('sheetIndex', String(sheetIndex));
  data.set('rowOffset', String(rowOffset));
  data.set('columnOffset', String(columnOffset));
  try {
    const response = await fetch('/api/excel/preview', {
      method: 'POST',
      body: data,
      signal: AbortSignal.any([request.signal, AbortSignal.timeout(120_000)])
    });
    const result = await response.json();
    if (request !== previewRequest) return;
    if (!response.ok) throw new Error(result.message || 'Preview unavailable.');
    preview = result;
    previewSheet.replaceChildren(...result.sheetNames.map((name, index) => new Option(name, index)));
    previewSheet.value = String(result.sheetIndex);
    $('preview-grid').scrollTop = 0;
    $('preview-grid').scrollLeft = 0;
    previewMessage = '';
  } catch (cause) {
    if (request !== previewRequest) return;
    previewFailed = true;
    previewMessage = cause.name === 'TimeoutError' ? 'Worksheet preview timed out.'
      : cause instanceof TypeError ? 'Could not reach the cleaner for a preview.'
      : 'Could not preview the worksheet. ' + cause.message;
  } finally {
    if (request === previewRequest) {
      previewRequest = null;
      renderPreview();
    }
  }
}

function tableCell(tag, text, scope) {
  const cell = document.createElement(tag);
  cell.textContent = text;
  if (scope) cell.scope = scope;
  return cell;
}

function renderPreview() {
  updateSubmit();
  $('preview').hidden = !file;
  $('preview-sheet-control').hidden = previewSheet.options.length < 2;
  $('preview-retry').hidden = !previewFailed;
  $('preview-status').textContent = previewMessage;
  const showPage = preview && !previewRequest && !previewFailed;
  $('preview-grid').hidden = !showPage || !preview.rows.length;
  const canPage = preview && (preview.rowOffset || preview.columnOffset || preview.moreRows || preview.moreColumns);
  $('preview-navigation').hidden = !canPage;
  for (const [id, available] of previewNavigation) {
    $(id).disabled = busy || !showPage || !available(preview);
  }
  if (!showPage) return;
  if (!preview.rows.length) {
    $('preview-status').textContent = 'This worksheet has no populated rows.';
    return;
  }
  const auto = headerMode.value === 'auto';
  const count = auto ? preview.detectedHeaderRows : headerRows.valueAsNumber;
  const description = Number.isInteger(count) && count >= 0 && count <= 100
    ? (count === 0 ? (auto ? 'No headers detected' : 'No header rows')
      : count + ' header ' + (count === 1 ? 'row' : 'rows') + (auto ? ' detected' : ' selected'))
    : 'Enter a header count from 0 to 100';
  const first = preview.firstPopulatedRow + preview.rowOffset;
  const last = first + preview.rows.length - 1;
  const columns = preview.columns[0] + '–' + preview.columns.at(-1);
  const range = 'Rows ' + first + '–' + last + ' of ' + (preview.firstPopulatedRow + preview.totalRows - 1)
    + ' · Columns ' + columns + ' (' + preview.totalColumns + ' in use)';
  $('preview-status').textContent = mode.value === 'columns'
    ? description + ' · First populated row: ' + preview.firstPopulatedRow + ' · ' + range
    : range;
  const table = $('preview-table');
  table.replaceChildren();
  table.createTHead().insertRow().append(tableCell('th', 'Row', 'col'),
    ...preview.columns.map(column => tableCell('th', column, 'col')));
  const body = table.createTBody();
  for (const entry of preview.rows) {
    const row = body.insertRow();
    const number = tableCell('th', entry.rowNumber, 'row');
    if (mode.value === 'columns' && entry.rowNumber < preview.firstPopulatedRow + count) {
      row.className = 'preview-header';
      const label = document.createElement('span');
      label.className = 'preview-row-kind';
      label.textContent = 'Header';
      number.append(label);
    }
    row.append(number, ...entry.values.map(value => tableCell('td', value)));
  }
}

function error(message = '') {
  $('error').textContent = message;
  $('error').hidden = !message;
}

function size(bytes) {
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function clearResult() {
  if (downloadUrl) URL.revokeObjectURL(downloadUrl);
  downloadUrl = null;
  $('download').removeAttribute('href');
  $('download').removeAttribute('download');
  $('result').hidden = true;
}

function renderSelection() {
  dropzone.hidden = Boolean(file);
  $('selected').hidden = !file;
  $('file-name').textContent = file?.name ?? '';
  $('file-info').textContent = file ? size(file.size) + (busy ? ' · Processing' : ' · Selected') : '';
  updateSubmit();
}

function updateSubmit() {
  $('submit').disabled = !file || busy || (mode.value === 'columns' && Boolean(previewRequest));
}

function choose(files) {
  if (busy || files.length === 0) return;
  const next = files[0];
  if (files.length !== 1) return error('Choose one workbook at a time.');
  if (!accepted.test(next.name)) return error('Choose an Excel workbook ending in .xlsx or .xls.');
  if (next.size === 0) return error('The uploaded file is empty.');
  if (next.size > 10 * 1024 * 1024) return error('Choose a workbook up to 10 MB.');
  file = next;
  clearPreview();
  error();
  clearResult();
  renderSelection();
  loadPreview();
  $('remove').focus();
}

function reset(defaults = false) {
  if (busy) return;
  file = null;
  clearPreview();
  input.value = '';
  if (defaults) {
    form.reset();
    renderMode();
  }
  error();
  clearResult();
  renderSelection();
  renderPreview();
  $('browse').focus();
}

function setBusy(value) {
  busy = value;
  form.setAttribute('aria-busy', String(value));
  $('browse').disabled = value;
  $('remove').disabled = value;
  mode.disabled = value;
  for (const option of options) option.disabled = value;
  renderMode();
  renderPreview();
  $('processing').hidden = !value;
  $('submit').querySelector('span').textContent = value ? 'Cleaning workbook…' : 'Clean Excel File';
  dropzone.setAttribute('aria-disabled', String(value));
  renderSelection();
}

function isFileDrag(event) {
  return Array.from(event.dataTransfer?.types ?? []).includes('Files');
}

dropzone.addEventListener('click', () => { if (!busy) input.click(); });
input.addEventListener('change', () => {
  choose(input.files);
  input.value = '';
});

window.addEventListener('dragover', (event) => {
  if (!isFileDrag(event)) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = !busy && form.contains(event.target) ? 'copy' : 'none';
});
window.addEventListener('drop', (event) => {
  if (!isFileDrag(event)) return;
  event.preventDefault();
  dragDepth = 0;
  dropzone.classList.remove('dragging');
  if (form.contains(event.target)) choose(event.dataTransfer.files);
});
dropzone.addEventListener('dragenter', (event) => {
  if (busy || !isFileDrag(event)) return;
  dragDepth++;
  dropzone.classList.add('dragging');
});
dropzone.addEventListener('dragleave', () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (!dragDepth) dropzone.classList.remove('dragging');
});
$('remove').addEventListener('click', () => reset());
$('reset').addEventListener('click', () => reset(true));

function resultName(disposition, fallback) {
  const encoded = disposition?.match(/filename\*=UTF-8''([^;]+)/i);
  if (!encoded) return fallback;
  try { return decodeURIComponent(encoded[1]); } catch { return fallback; }
}

form.addEventListener('submit', async (event) => {
  event.preventDefault();
  if ($('submit').disabled) return;
  error();
  clearResult();
  const data = new FormData();
  const columns = mode.value === 'columns';
  data.set('file', file);
  data.set('removeEmptyColumns', String(columns));
  if (columns) {
    data.set('headerRows', headerMode.value === 'auto' ? '-1' : headerRows.value);
  } else {
    for (const option of options) data.set(option.name, String(option.checked));
  }
  setBusy(true);
  let complete = false;
  try {
    const response = await fetch('/api/excel/clean', {
      method: 'POST',
      body: data,
      signal: AbortSignal.timeout(120_000)
    });
    if (!response.ok) {
      const problem = await response.json().catch(() => null);
      error(problem?.message || 'The workbook could not be cleaned. Try again.');
      return;
    }
    const headers = response.headers;
    const blob = await response.blob();
    const filename = resultName(headers.get('content-disposition'),
      file.name.replace(accepted, (extension) => '_cleaned' + extension.toLowerCase()));
    const sheets = Number(headers.get('x-worksheets-processed'));
    const changes = Number(headers.get(columns ? 'x-columns-removed' : 'x-cells-cleared'));
    const milliseconds = Number(headers.get('x-processing-millis'));
    $('original').textContent = file.name;
    $('cleaned').textContent = filename;
    $('sheets').textContent = sheets.toLocaleString();
    $('changes-label').textContent = columns ? 'Columns removed' : 'Cells cleared';
    $('changes').textContent = changes.toLocaleString();
    $('duration').textContent = milliseconds < 1000 ? milliseconds + ' ms' : (milliseconds / 1000).toFixed(2) + ' s';
    const noChanges = columns ? 'No empty columns found for removal.' : 'No matching values found.';
    $('result-summary').textContent = changes ? 'Your file is ready to download.' : noChanges;
    downloadUrl = URL.createObjectURL(blob);
    $('download').href = downloadUrl;
    $('download').download = filename;
    $('result').hidden = false;
    complete = true;
  } catch (cause) {
    error(cause.name === 'TimeoutError'
      ? 'Cleaning took too long. Try a smaller workbook.'
      : 'Could not reach the cleaner. Check that the app is running and try again.');
  } finally {
    setBusy(false);
    if (complete) $('result-heading').focus();
    else $('submit').focus();
  }
});
