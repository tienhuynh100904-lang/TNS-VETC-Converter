import * as pdfjsLib from 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.7.76/pdf.min.mjs';

pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.7.76/pdf.worker.min.mjs';

// Some VETC invoice PDFs (e.g. ones not going through the "chuyển đổi từ
// hóa đơn điện tử" conversion path) are rendered with one text item per
// glyph instead of one item per word — including splitting accented
// Vietnamese letters into their own item ("K","ý" instead of "Ký"). Those
// glyph items sit flush against each other on the X axis with no real gap,
// so we must only insert a space when there is an actual visual gap
// between items — never unconditionally between every item — or accented
// words get torn apart ("Ký hiệu" -> "K ý  hi ệ u") and every downstream
// regex silently stops matching.
function joinTextItems(textContent) {
  let out = '';
  let lastY = null;
  let lastX = null;
  for (const item of textContent.items) {
    if (item.str === '') continue; // invisible position markers, no width to measure from
    const y = item.transform[5];
    const x = item.transform[4];
    const newLine = lastY !== null && Math.abs(y - lastY) > 1;
    if (newLine) {
      out += '\n';
    } else if (lastX !== null) {
      const gap = x - lastX;
      // Threshold scales with glyph height so it works across the
      // document's different font sizes (title vs. body text).
      const threshold = Math.max(item.height || 0, 1) * 0.28;
      if (gap > threshold && !out.endsWith(' ')) out += ' ';
    }
    out += item.str;
    lastY = y;
    lastX = x + (item.width || 0);
  }
  return out;
}

// Browser memory guard for large batches: a VETC invoice PDF is normally a
// few hundred KB (largest sample ~185KB). 15MB comfortably covers a
// high-res scanned page while still catching corrupt/oversized files before
// they bloat memory across a multi-thousand-file batch.
export const MAX_FILE_SIZE_BYTES = 15 * 1024 * 1024;

async function extractText(file) {
  if (file.size > MAX_FILE_SIZE_BYTES) {
    throw new Error(`File quá lớn (${(file.size / 1024 / 1024).toFixed(1)}MB, giới hạn ${MAX_FILE_SIZE_BYTES / 1024 / 1024}MB)`);
  }
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;
  // Each page is text-joined separately (not concatenated as one blob)
  // so a multi-page PDF can be checked per page below — e.g. to tell a
  // genuinely repeated toll-pass line apart from the same invoice page
  // being duplicated (a reprinted copy on page 2).
  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    pages.push(joinTextItems(content));
  }
  return { text: pages.join('\n'), pageCount: pdf.numPages };
}

function toNumber(str) {
  if (!str) return 0;
  const cleaned = str.replace(/\./g, '').replace(/,/g, '.').replace(/[^\d.-]/g, '');
  const n = parseFloat(cleaned);
  return Number.isFinite(n) ? n : 0;
}

function pad2(n) { return String(n).padStart(2, '0'); }

/**
 * Parse plain text of a VETC "Hóa đơn giá trị gia tăng" invoice into a flat row
 * matching the columns of the ETCT8.xlsx sample statement.
 */
function parseInvoiceText(text, fileName) {
  // PDF.js emits real newlines between visually-stacked fragments (e.g. a
  // wrapped transaction description, or a value pushed to the next line
  // after a label). Normalize label:value gaps but keep newlines available
  // for patterns that need to span them.
  const clean = text.replace(/\r/g, '');

  // NOTE: use [ \t]* (not \s*) right after ":" for same-line values — \s*
  // would also swallow the newline and let the match spill onto the next
  // label's line when the value is empty (e.g. buyer's "Mã số thuế :" is
  // often blank, immediately followed by "Địa chỉ :...").
  const symbolMatch = clean.match(/Ký hiệu[ \t]*:[ \t]*(\S+)/);
  const rawSymbol = symbolMatch ? symbolMatch[1].trim() : '';
  // Sample export drops the leading serial digit ("1K26TEF" -> "K26TEF")
  const symbol = rawSymbol.replace(/^\d/, '');

  const invoiceNoMatch = clean.match(/(?:^|\n)Số[ \t]*:[ \t]*\n?[ \t]*(\d+)/);
  const invoiceNoRaw = invoiceNoMatch ? invoiceNoMatch[1] : '';
  // Every sample invoice number is 7-8 digits. A match outside a sane range
  // (6-12 digits) more likely means the regex grabbed the wrong number than
  // a real invoice number — treat it as not found rather than trust it.
  const invoiceNo = /^\d{6,12}$/.test(invoiceNoRaw) ? invoiceNoRaw : '';

  const invoiceDateMatch = clean.match(/Ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})/);
  let invoiceDate = '';
  if (invoiceDateMatch) {
    const day = Number(invoiceDateMatch[1]);
    const month = Number(invoiceDateMatch[2]);
    const year = Number(invoiceDateMatch[3]);
    // Guard against an out-of-range day/month silently rolling over into a
    // different date (e.g. JS Date turns "32/13" into some other real date
    // without complaint) — treat it as unparsed instead.
    const isValidCalendarDate = month >= 1 && month <= 12 && day >= 1 && day <= 31;
    if (isValidCalendarDate) {
      invoiceDate = `${pad2(day)}/${pad2(month)}/${year}`;
    }
  }

  const buyerNameMatch = clean.match(/Họ tên người mua hàng[ \t]*:[ \t]*([^\n]+)/);
  const buyerName = buyerNameMatch ? buyerNameMatch[1].trim() : '';

  let buyerAddress = '';
  {
    // second "Địa chỉ" occurrence belongs to the buyer block, first is the seller
    const addrMatches = [...clean.matchAll(/Địa chỉ[ \t]*:[ \t]*([^\n]*)/g)];
    if (addrMatches.length >= 2) buyerAddress = addrMatches[1][1].trim();
  }

  const taxCodeMatches = [...clean.matchAll(/Mã số thuế[ \t]*:[ \t]*([^\n]*)/g)];
  const buyerTaxCode = taxCodeMatches.length >= 2 ? taxCodeMatches[1][1].trim() : '';

  // Core transaction line(s), e.g.:
  // "Cước đường bộ xe 51M95010T đi qua trạm\nChơn Thành thời gian GD 12:19:42 ngày\n01/08/2026 mã GD 3278914483"
  // The description can wrap across several real newlines, so match across
  // whitespace of any kind ([\s\S]) rather than assuming a single line.
  // Every sample invoice on file has exactly one such line (one toll pass per
  // invoice). We only support that case — if a PDF ever contains more than
  // one, we deliberately do NOT guess how to merge them; the invoice is
  // flagged as an error so a human reviews it instead of silently mis-mapping
  // amounts/content to the wrong transaction.
  const txPattern = /Cước đường bộ xe\s+(\S+)\s+đi qua\s+([\s\S]+?)\s+thời gian GD\s+(\d{1,2}:\d{2}:\d{2})\s+ngày\s+(\d{2}\/\d{2}\/\d{4})\s+mã GD\s+(-?\d+)/g;
  const allTxMatches = [...clean.matchAll(txPattern)];
  // A multi-page PDF can legitimately repeat the exact same invoice content
  // (e.g. a duplicated page, a reprinted copy). That is not "multiple line
  // items" — dedupe by transaction code (group 5), which is unique per real
  // toll pass, before deciding whether this invoice genuinely has more than
  // one line item.
  const seenTxCodes = new Set();
  const txMatches = allTxMatches.filter((m) => {
    if (seenTxCodes.has(m[5])) return false;
    seenTxCodes.add(m[5]);
    return true;
  });
  const txMatch = txMatches[0];

  let plateNumber = '';
  let station = '';
  let content = '';
  let transactionCode = '';
  let receiptDateTime = '';

  if (txMatch) {
    plateNumber = txMatch[1];
    station = txMatch[2].replace(/\s+/g, ' ').replace(/^(trạm|đoạn)\s+/i, '').trim();
    const time = txMatch[3];
    const date = txMatch[4];
    transactionCode = txMatch[5];
    receiptDateTime = `${date} ${time}`;
    content = txMatch[0].replace(/\s+/g, ' ').trim();
  }

  const totalMatch = clean.match(
    /Tổng tiền\s*:\s*(-?[\d.,]+)\s+(-?[\d.,]+)\s+(-?[\d.,]+)/
  );
  const amountBeforeTax = totalMatch ? toNumber(totalMatch[1]) : 0;
  const vatAmount = totalMatch ? toNumber(totalMatch[2]) : 0;
  const amountAfterTax = totalMatch ? toNumber(totalMatch[3]) : 0;

  // Cross-check: with a single transaction line, its own unit-price columns
  // (5)=Đơn giá and (7)=Thành tiền chưa thuế must equal the "Tổng tiền" row,
  // since there is nothing else to sum. A mismatch means our column-reading
  // assumption broke for this layout — surface it instead of trusting it.
  let lineAmountMatch = null;
  if (txMatch) {
    const afterTx = clean.slice(txMatch.index + txMatch[0].length);
    lineAmountMatch = afterTx.match(
      /^[\s\S]{0,40}?(-?[\d.,]+)\s+(-?[\d.,]+)\s*%\s*(-?[\d.,]+)\s+(-?[\d.,]+)\s+(-?[\d.,]+)/
    );
  }
  const amountMismatch =
    txMatch && totalMatch && lineAmountMatch &&
    toNumber(lineAmountMatch[3]) !== amountBeforeTax;

  const row = {
    symbol,
    docType: 'Hóa đơn',
    invoiceNo,
    invoiceDate,
    receiptDateTime,
    transactionCode,
    plateNumber,
    content,
    amountBeforeTax,
    vatAmount,
    amountAfterTax,
    station,
    invoiceKind: 'Hóa đơn gốc',
    status: 'Đã duyệt',
    fileName,
    buyerName,
    buyerAddress,
    buyerTaxCode,
  };

  const missing = [];
  if (!invoiceNo) missing.push('Số hóa đơn');
  if (!invoiceDate) missing.push('Ngày hóa đơn');
  if (!txMatch) missing.push('Nội dung giao dịch');
  if (!totalMatch) missing.push('Số tiền');
  if (txMatches.length > 1) {
    missing.push(`Hóa đơn có ${txMatches.length} dòng hàng hóa (chỉ hỗ trợ 1 dòng/hóa đơn)`);
  }
  if (amountMismatch) {
    missing.push('Số tiền dòng hàng hóa không khớp với dòng Tổng tiền');
  }

  return { row, missing };
}

export async function parsePdfInvoice(file) {
  const { text, pageCount } = await extractText(file);
  const { row, missing } = parseInvoiceText(text, file.name);
  row.pageCount = pageCount;
  if (missing.length > 0) {
    throw Object.assign(
      new Error(`Thiếu dữ liệu: ${missing.join(', ')}`),
      { row, missing }
    );
  }
  return row;
}
