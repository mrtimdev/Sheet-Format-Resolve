'use strict';

const $ = (id) => document.getElementById(id);
const form = $('cleaner');
const dropzone = $('dropzone');
const input = $('file');
const options = [...form.querySelectorAll('input[type="checkbox"]')];
const accepted = /\.(xlsx|xls)$/i;
let file = null;
let busy = false;
let downloadUrl = null;
let dragDepth = 0;

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
  $('submit').disabled = !file || busy;
}

function choose(files) {
  if (busy || files.length === 0) return;
  const next = files[0];
  if (files.length !== 1) return error('Choose one workbook at a time.');
  if (!accepted.test(next.name)) return error('Choose an Excel workbook ending in .xlsx or .xls.');
  if (next.size === 0) return error('The uploaded file is empty.');
  if (next.size > 10 * 1024 * 1024) return error('Choose a workbook up to 10 MB.');
  file = next;
  error();
  clearResult();
  renderSelection();
  $('remove').focus();
}

function reset(defaults = false) {
  if (busy) return;
  file = null;
  input.value = '';
  if (defaults) for (const option of options) option.checked = option.defaultChecked;
  error();
  clearResult();
  renderSelection();
  $('browse').focus();
}

function setBusy(value) {
  busy = value;
  form.setAttribute('aria-busy', String(value));
  $('browse').disabled = value;
  $('remove').disabled = value;
  for (const option of options) option.disabled = value;
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
  if (!file || busy) return;
  error();
  clearResult();
  const data = new FormData();
  data.set('file', file);
  for (const option of options) data.set(option.name, String(option.checked));
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
    const cells = Number(headers.get('x-cells-cleared'));
    const milliseconds = Number(headers.get('x-processing-millis'));
    $('original').textContent = file.name;
    $('cleaned').textContent = filename;
    $('sheets').textContent = sheets.toLocaleString();
    $('cells').textContent = cells.toLocaleString();
    $('duration').textContent = milliseconds < 1000 ? milliseconds + ' ms' : (milliseconds / 1000).toFixed(2) + ' s';
    $('result-summary').textContent = cells ? 'Your file is ready to download.' : 'No matching values found. Your workbook is unchanged.';
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
