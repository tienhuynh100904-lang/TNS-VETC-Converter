import { parsePdfInvoice, MAX_FILE_SIZE_BYTES } from './pdf-parser.js';
import { downloadWorkbook } from './excel-builder.js';

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const folderInput = document.getElementById('folderInput');
const btnPickFiles = document.getElementById('btnPickFiles');
const btnPickFolder = document.getElementById('btnPickFolder');
const btnExport = document.getElementById('btnExport');
const btnClear = document.getElementById('btnClear');

const statusPanel = document.getElementById('statusPanel');
const statusText = document.getElementById('statusText');
const countTotal = document.getElementById('countTotal');
const countOk = document.getElementById('countOk');
const countErr = document.getElementById('countErr');
const progressFill = document.getElementById('progressFill');

const errorsPanel = document.getElementById('errorsPanel');
const errorList = document.getElementById('errorList');
const errCount = document.getElementById('errCount');

const duplicatesPanel = document.getElementById('duplicatesPanel');
const duplicateList = document.getElementById('duplicateList');
const dupCount = document.getElementById('dupCount');

const summaryPanel = document.getElementById('summaryPanel');
const sumInvoiceCount = document.getElementById('sumInvoiceCount');
const sumBeforeTax = document.getElementById('sumBeforeTax');
const sumVat = document.getElementById('sumVat');
const sumAfterTax = document.getElementById('sumAfterTax');

const tablePanel = document.getElementById('tablePanel');
const dataTableBody = document.getElementById('dataTableBody');
const rowCount = document.getElementById('rowCount');
const searchBox = document.getElementById('searchBox');
const sizeLimitHint = document.getElementById('sizeLimitHint');

const CONCURRENCY = 6;

sizeLimitHint.textContent = `Giới hạn mỗi file: ${Math.round(MAX_FILE_SIZE_BYTES / 1024 / 1024)}MB — mỗi hóa đơn phải chỉ có 1 dòng hàng hóa (1 lượt qua trạm); hóa đơn nhiều dòng sẽ bị đánh dấu lỗi để bạn kiểm tra lại thủ công`;

/** @type {Array<object>} */
let successRows = [];
/** @type {Array<{fileName: string, message: string}>} */
let errorRows = [];
/** @type {Array<{fileName: string, keptFileName: string}>} */
let duplicateRows = [];
let processing = false;

const LARGE_BATCH_WARNING_THRESHOLD = 3000;

function resetState() {
  successRows = [];
  errorRows = [];
  duplicateRows = [];
  dataTableBody.innerHTML = '';
  errorList.innerHTML = '';
  statusPanel.hidden = true;
  summaryPanel.hidden = true;
  errorsPanel.hidden = true;
  duplicatesPanel.hidden = true;
  duplicateList.innerHTML = '';
  tablePanel.hidden = true;
  btnExport.disabled = true;
  progressFill.style.width = '0%';
}

function isPdfFile(file) {
  return file.type === 'application/pdf' || /\.pdf$/i.test(file.name);
}

function collectPdfFiles(fileList) {
  return Array.from(fileList).filter(isPdfFile);
}

const ROW_FIELDS = [
  'symbol', 'docType', 'invoiceNo', 'invoiceDate', 'receiptDateTime',
  'transactionCode', 'plateNumber', 'content', 'amountBeforeTax',
  'vatAmount', 'amountAfterTax', 'station', 'invoiceKind', 'status',
];

function appendRowToTable(row, index) {
  const tr = document.createElement('tr');
  tr.dataset.index = String(index);

  const sttTd = document.createElement('td');
  sttTd.textContent = String(index + 1);
  tr.appendChild(sttTd);

  const buyerTd = document.createElement('td');
  buyerTd.contentEditable = 'true';
  buyerTd.dataset.field = 'buyerName';
  buyerTd.textContent = row.buyerName || '';
  buyerTd.addEventListener('blur', () => onCellEdit(buyerTd, index, 'buyerName'));
  tr.appendChild(buyerTd);

  ROW_FIELDS.forEach((field) => {
    const td = document.createElement('td');
    td.contentEditable = 'true';
    td.dataset.field = field;
    const value = row[field];
    td.textContent = typeof value === 'number' ? formatNumber(value) : (value || '');
    td.addEventListener('blur', () => onCellEdit(td, index, field));
    tr.appendChild(td);
  });

  const fileTd = document.createElement('td');
  fileTd.textContent = row.fileName;
  fileTd.style.color = '#9ca3af';
  fileTd.style.fontSize = '12px';
  tr.appendChild(fileTd);

  dataTableBody.appendChild(tr);
}

function formatNumber(n) {
  return new Intl.NumberFormat('vi-VN').format(n);
}

const REQUIRED_FIELDS = ['invoiceNo', 'invoiceDate'];

function onCellEdit(td, index, field) {
  const raw = td.textContent.trim();
  const isMoney = ['amountBeforeTax', 'vatAmount', 'amountAfterTax'].includes(field);
  if (isMoney) {
    const n = Number(raw.replace(/\./g, '').replace(/,/g, '.').replace(/[^\d.-]/g, ''));
    successRows[index][field] = Number.isFinite(n) ? n : 0;
    td.textContent = formatNumber(successRows[index][field]);
  } else {
    successRows[index][field] = raw;
  }
  td.classList.toggle('cell-error', REQUIRED_FIELDS.includes(field) && raw === '');
  if (isMoney) updateSummary();
}

function updateCounts() {
  countTotal.textContent = String(successRows.length + errorRows.length + duplicateRows.length);
  countOk.textContent = String(successRows.length);
  countErr.textContent = String(errorRows.length);
  rowCount.textContent = String(successRows.length);
  updateSummary();
}

function updateSummary() {
  summaryPanel.hidden = successRows.length === 0;
  const totals = successRows.reduce((acc, row) => {
    acc.before += Number(row.amountBeforeTax) || 0;
    acc.vat += Number(row.vatAmount) || 0;
    acc.after += Number(row.amountAfterTax) || 0;
    return acc;
  }, { before: 0, vat: 0, after: 0 });

  sumInvoiceCount.textContent = formatNumber(successRows.length);
  sumBeforeTax.textContent = formatNumber(totals.before) + ' đ';
  sumVat.textContent = formatNumber(totals.vat) + ' đ';
  sumAfterTax.textContent = formatNumber(totals.after) + ' đ';
}

function renderErrors() {
  errorList.innerHTML = '';
  errCount.textContent = String(errorRows.length);
  errorsPanel.hidden = errorRows.length === 0;
  errorRows.forEach((e) => {
    const li = document.createElement('li');
    li.innerHTML = `<b>${e.fileName}</b> — ${e.message}. <i>Vui lòng mở lại file này để kiểm tra thủ công.</i>`;
    errorList.appendChild(li);
  });
}

function renderDuplicates() {
  duplicateList.innerHTML = '';
  dupCount.textContent = String(duplicateRows.length);
  duplicatesPanel.hidden = duplicateRows.length === 0;
  duplicateRows.forEach((d) => {
    const li = document.createElement('li');
    li.innerHTML = `<b>${d.fileName}</b> — cùng hóa đơn với <b>${d.keptFileName}</b>, đã loại bỏ để tránh tính trùng.`;
    duplicateList.appendChild(li);
  });
}

async function runQueue(files) {
  let cursor = 0;
  let done = 0;
  const total = files.length;
  // Tracks (symbol + invoiceNo) -> file name that already claimed it, so two
  // files describing the same real invoice (e.g. accidentally dropped twice)
  // don't both land in the export and double-count the money.
  const seenInvoiceKeys = new Map();

  async function worker() {
    while (cursor < files.length) {
      const idx = cursor++;
      const file = files[idx];
      try {
        const row = await parsePdfInvoice(file);
        // Everything from here to the next await is synchronous, so no
        // other worker can interleave between the check and the claim.
        const dupKey = `${row.symbol}|||${row.invoiceNo}`;
        const existingFile = seenInvoiceKeys.get(dupKey);
        if (existingFile) {
          duplicateRows.push({ fileName: file.name, keptFileName: existingFile });
        } else {
          seenInvoiceKeys.set(dupKey, file.name);
          successRows.push(row);
          appendRowToTable(row, successRows.length - 1);
        }
      } catch (err) {
        errorRows.push({ fileName: file.name, message: err.message || 'Lỗi không xác định' });
      }
      done++;
      const pct = Math.round((done / total) * 100);
      progressFill.style.width = pct + '%';
      statusText.textContent = `Đang xử lý ${done}/${total} file…`;
      updateCounts();
    }
  }

  const workers = Array.from({ length: Math.min(CONCURRENCY, total) }, () => worker());
  await Promise.all(workers);
}

async function handleFiles(fileList) {
  const pdfFiles = collectPdfFiles(fileList);
  if (pdfFiles.length === 0) {
    alert('Không tìm thấy file PDF nào trong lựa chọn của bạn.');
    return;
  }
  if (processing) return;

  if (pdfFiles.length > LARGE_BATCH_WARNING_THRESHOLD) {
    const proceed = confirm(
      `Bạn đang xử lý ${pdfFiles.length.toLocaleString('vi-VN')} file PDF.\n` +
      `Quá trình này có thể mất vài phút và tiêu tốn nhiều bộ nhớ trình duyệt.\n\nTiếp tục?`
    );
    if (!proceed) return;
  }

  processing = true;

  resetState();
  statusPanel.hidden = false;
  tablePanel.hidden = false;
  updateCounts();
  countTotal.textContent = String(pdfFiles.length);
  statusText.textContent = `Đang xử lý 0/${pdfFiles.length} file…`;

  await runQueue(pdfFiles);

  processing = false;
  const buyerCount = new Set(successRows.map((r) => `${r.buyerName || ''}|||${r.buyerTaxCode || ''}`)).size;
  const buyerNote = buyerCount > 1 ? ` — phát hiện ${buyerCount} người mua hàng khác nhau, sẽ xuất thành ${buyerCount} sheet riêng.` : '';
  const dupNote = duplicateRows.length > 0 ? ` — đã loại bỏ ${duplicateRows.length} hóa đơn trùng lặp.` : '';
  if (successRows.length === 0) {
    statusText.textContent = `Không có hóa đơn nào xử lý thành công trong ${pdfFiles.length} file đã chọn. Vui lòng kiểm tra danh sách lỗi bên dưới.`;
  } else {
    statusText.textContent = `Hoàn tất: ${successRows.length}/${pdfFiles.length} file xử lý thành công.${buyerNote}${dupNote}`;
  }
  renderErrors();
  renderDuplicates();
  btnExport.disabled = successRows.length === 0;
}

// --- Event wiring ---

btnPickFiles.addEventListener('click', () => fileInput.click());
btnPickFolder.addEventListener('click', () => folderInput.click());

fileInput.addEventListener('change', (e) => {
  handleFiles(e.target.files);
  fileInput.value = '';
});

folderInput.addEventListener('change', (e) => {
  handleFiles(e.target.files);
  folderInput.value = '';
});

['dragenter', 'dragover'].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.add('dragover');
  });
});

['dragleave', 'drop'].forEach((evt) => {
  dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropzone.classList.remove('dragover');
  });
});

dropzone.addEventListener('drop', async (e) => {
  const items = e.dataTransfer.items;
  if (items && items.length && items[0].webkitGetAsEntry) {
    const files = await readDroppedEntries(items);
    handleFiles(files);
  } else {
    handleFiles(e.dataTransfer.files);
  }
});

async function readDroppedEntries(items) {
  const entries = Array.from(items)
    .map((it) => it.webkitGetAsEntry())
    .filter(Boolean);
  const files = [];

  async function walk(entry) {
    if (entry.isFile) {
      const file = await new Promise((resolve, reject) => entry.file(resolve, reject));
      files.push(file);
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      const readBatch = () => new Promise((resolve, reject) => reader.readEntries(resolve, reject));
      let batch;
      do {
        batch = await readBatch();
        for (const child of batch) await walk(child);
      } while (batch.length > 0);
    }
  }

  for (const entry of entries) await walk(entry);
  return files;
}

btnExport.addEventListener('click', () => {
  if (successRows.length === 0) return;
  downloadWorkbook(successRows);
});

btnClear.addEventListener('click', () => {
  resetState();
});

searchBox.addEventListener('input', () => {
  const q = searchBox.value.trim().toLowerCase();
  Array.from(dataTableBody.rows).forEach((tr) => {
    const text = tr.textContent.toLowerCase();
    tr.style.display = text.includes(q) ? '' : 'none';
  });
});
