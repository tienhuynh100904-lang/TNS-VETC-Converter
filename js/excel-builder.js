/* global XLSX */

const HEADERS = [
  'STT',
  'Ký hiệu hoá đơn',
  'Loại chứng từ điện tử',
  'Số hóa đơn',
  'Ngày phát sinh hóa đơn',
  'Ngày phát sinh biên lai',
  'Mã giao dịch',
  'Biển số xe',
  'Nội dung hoá đơn',
  'Số tiền chưa thuế',
  'Số tiền thuế GTGT',
  'Số tiền sau thuế',
  'Trạm',
  'Loại hoá đơn',
  'Trạng thái hóa đơn',
];

const COL_WIDTHS = [5, 12, 14, 12, 16, 17, 12, 12, 46, 14, 14, 14, 20, 12, 14];

function parseDdMmYyyy(str) {
  const m = /^(\d{2})\/(\d{2})\/(\d{4})/.exec(str || '');
  if (!m) return null;
  return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
}

function formatDateRangeLabel(rows) {
  const dates = rows
    .map((r) => parseDdMmYyyy(r.invoiceDate))
    .filter(Boolean)
    .sort((a, b) => a - b);
  if (dates.length === 0) return '';
  const first = dates[0];
  const last = dates[dates.length - 1];
  const fmt = (d) => `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`;
  return `Từ ngày ${fmt(first)} đến hết ngày ${fmt(last)}`;
}

/**
 * Group invoice rows by buyer. Rows with no buyer name at all are grouped
 * together under an "(Không xác định)" bucket rather than silently merged
 * into whichever buyer happens to be first.
 */
function groupByBuyer(rows) {
  const groups = new Map();
  rows.forEach((r) => {
    const key = `${r.buyerName || ''}|||${r.buyerTaxCode || ''}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(r);
  });
  return [...groups.values()];
}

function sanitizeSheetName(name, usedNames) {
  let base = (name || 'Người mua').replace(/[:\\/?*[\]]/g, ' ').trim().slice(0, 31) || 'Người mua';
  let candidate = base;
  let n = 2;
  while (usedNames.has(candidate)) {
    const suffix = ` (${n})`;
    candidate = base.slice(0, 31 - suffix.length) + suffix;
    n++;
  }
  usedNames.add(candidate);
  return candidate;
}

/**
 * Build a single sheet matching the layout of the ETCT8.xlsx sample:
 * company header block, buyer info block, 15-column data table, totals row.
 */
function buildSheet(rows) {
  const first = rows[0] || {};
  const sellerName = 'CÔNG TY TNHH THU PHÍ TỰ ĐỘNG VETC';

  const aoa = [];
  aoa.push([sellerName]);
  aoa.push(['BẢNG KÊ HOÁ ĐƠN ĐIỆN TỬ']);
  aoa.push([formatDateRangeLabel(rows)]);
  aoa.push([`Tên người mua hàng: ${first.buyerName || ''}`]);
  aoa.push([`Địa chỉ: ${first.buyerAddress || ''}`]);
  aoa.push([`Mã số thuế: ${first.buyerTaxCode || ''}`]);
  aoa.push([]);
  aoa.push(HEADERS);

  const headerRowIndex = aoa.length - 1;
  const dataStartRow = aoa.length;

  rows.forEach((r, i) => {
    aoa.push([
      i + 1,
      r.symbol,
      r.docType,
      r.invoiceNo,
      r.invoiceDate,
      r.receiptDateTime,
      r.transactionCode,
      r.plateNumber,
      r.content,
      r.amountBeforeTax,
      r.vatAmount,
      r.amountAfterTax,
      r.station,
      r.invoiceKind,
      r.status,
    ]);
  });

  const dataEndRow = aoa.length - 1;

  const totalRow = new Array(HEADERS.length).fill('');
  totalRow[8] = 'Tổng cộng:';
  aoa.push(totalRow);
  const totalRowIndex = aoa.length - 1;

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  // SUM formulas for the three money columns (J, K, L => 0-indexed 9,10,11)
  [9, 10, 11].forEach((col) => {
    const colLetter = XLSX.utils.encode_col(col);
    const cellRef = XLSX.utils.encode_cell({ r: totalRowIndex, c: col });
    ws[cellRef] = {
      t: 'n',
      f: `SUM(${colLetter}${dataStartRow + 1}:${colLetter}${dataEndRow + 1})`,
    };
  });

  ws['!cols'] = COL_WIDTHS.map((wch) => ({ wch }));

  ws['!merges'] = [
    { s: { r: 0, c: 0 }, e: { r: 0, c: HEADERS.length - 1 } },
    { s: { r: 1, c: 0 }, e: { r: 1, c: HEADERS.length - 1 } },
    { s: { r: 2, c: 0 }, e: { r: 2, c: HEADERS.length - 1 } },
    { s: { r: 3, c: 0 }, e: { r: 3, c: HEADERS.length - 1 } },
    { s: { r: 4, c: 0 }, e: { r: 4, c: HEADERS.length - 1 } },
    { s: { r: 5, c: 0 }, e: { r: 5, c: HEADERS.length - 1 } },
    { s: { r: totalRowIndex, c: 0 }, e: { r: totalRowIndex, c: 8 } },
  ];

  const numberFmt = '#,##0';
  for (let r = dataStartRow; r <= dataEndRow; r++) {
    [9, 10, 11].forEach((c) => {
      const ref = XLSX.utils.encode_cell({ r, c });
      if (ws[ref]) ws[ref].z = numberFmt;
    });
  }
  [9, 10, 11].forEach((c) => {
    const ref = XLSX.utils.encode_cell({ r: totalRowIndex, c });
    if (ws[ref]) ws[ref].z = numberFmt;
  });

  // Bold the title, header row and totals row
  const boldRefs = [
    XLSX.utils.encode_cell({ r: 0, c: 0 }),
    XLSX.utils.encode_cell({ r: 1, c: 0 }),
  ];
  for (let c = 0; c < HEADERS.length; c++) {
    boldRefs.push(XLSX.utils.encode_cell({ r: headerRowIndex, c }));
    boldRefs.push(XLSX.utils.encode_cell({ r: totalRowIndex, c }));
  }
  boldRefs.forEach((ref) => {
    if (ws[ref]) ws[ref].s = { font: { bold: true } };
  });

  return ws;
}

/**
 * Build a workbook with one sheet per distinct buyer (matched on name + tax
 * code), so a batch mixing invoices from several buyers doesn't mislabel
 * everyone under whichever buyer happened to be parsed first.
 */
export function buildWorkbook(rows) {
  const groups = groupByBuyer(rows);
  const wb = XLSX.utils.book_new();
  const usedNames = new Set();

  groups.forEach((groupRows) => {
    const ws = buildSheet(groupRows);
    const buyerName = groupRows[0]?.buyerName || 'Không xác định';
    const sheetName = sanitizeSheetName(buyerName, usedNames);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  });

  return wb;
}

export function downloadWorkbook(rows) {
  const wb = buildWorkbook(rows);
  const now = new Date();
  const stamp = [
    now.getFullYear(),
    String(now.getMonth() + 1).padStart(2, '0'),
    String(now.getDate()).padStart(2, '0'),
  ].join('') + '_' + [
    String(now.getHours()).padStart(2, '0'),
    String(now.getMinutes()).padStart(2, '0'),
  ].join('');
  XLSX.writeFile(wb, `TNS_VECT_Converter_${stamp}.xlsx`);
}
