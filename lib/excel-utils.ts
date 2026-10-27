import ExcelJS from 'exceljs';

export interface MergeConfig {
  includeTotal: boolean;
  startRow: number;
  startColumn: string;
  endColumn: string;
  includeSignature: boolean;
  outputFormat?: 'auto' | 'xlsx' | 'csv';
}

/**
 * Find the row number that contains "TOTAL" in any column
 * Similar to VBA logic - searches across all columns, returns first found
 */
export function findTotalRow(worksheet: ExcelJS.Worksheet): number | null {
  let totalRow: number | null = null;

  worksheet.eachRow((row, rowNumber) => {
    if (totalRow) return; // Exit early if already found

    let foundTotal = false;
    row.eachCell(cell => {
      const cellValue = cell.value?.toString().trim().toUpperCase();
      if (cellValue === 'TOTAL') {
        foundTotal = true;
      }
    });

    if (foundTotal) {
      totalRow = rowNumber;
    }
  });

  return totalRow;
}

/**
 * Check if a row is a subtotal/group total row
 * A subtotal row typically has:
 * - Most cells are empty
 * - Only 1-2 cells have numeric values (quantity/amount)
 * - No text content in identifier columns
 */
export function isSubtotalRow(row: ExcelJS.Row, startCol: number, endCol: number): boolean {
  let emptyCellCount = 0;
  let numericCellCount = 0;
  let textCellCount = 0;
  const totalCells = endCol - startCol + 1;

  for (let colIndex = startCol; colIndex <= endCol; colIndex++) {
    const cell = row.getCell(colIndex);
    const cellValue = cell.value;

    if (cellValue === null || cellValue === undefined || cellValue === '') {
      emptyCellCount++;
    } else if (typeof cellValue === 'number') {
      numericCellCount++;
    } else {
      const strValue = cellValue.toString().trim();
      if (strValue === '') {
        emptyCellCount++;
      } else if (!isNaN(Number(strValue.replace(/,/g, '')))) {
        // String that looks like a number (e.g., "424,595")
        numericCellCount++;
      } else {
        textCellCount++;
      }
    }
  }

  // A subtotal row has mostly empty cells and only 1-3 numeric values, no meaningful text
  // At least 70% of cells should be empty and very few text cells
  const emptyRatio = emptyCellCount / totalCells;
  return emptyRatio >= 0.7 && numericCellCount >= 1 && numericCellCount <= 3 && textCellCount === 0;
}

/**
 * Find signature row - looks for signature keywords
 * Supports: "Người lập phiếu", "Người nhận", "Signature of sender", "Chữ ký người gửi"
 */
export function findSignatureRow(worksheet: ExcelJS.Worksheet): number | null {
  let signatureRow: number | null = null;

  worksheet.eachRow((row, rowNumber) => {
    if (signatureRow) return; // Already found

    row.eachCell(cell => {
      const cellValue = cell.value?.toString().trim().toLowerCase();
      if (!cellValue) return;

      // Check for signature keywords
      const signatureKeywords = ['người lập phiếu', 'người nhận', 'signature of sender', 'chữ ký người gửi', 'signature', 'chữ ký'];

      for (const keyword of signatureKeywords) {
        if (cellValue.includes(keyword) && !signatureRow) {
          signatureRow = rowNumber;
          return;
        }
      }
    });
  });

  return signatureRow;
}

/**
 * Convert column letter (A, B, C) to column number (1, 2, 3)
 */
export function columnLetterToNumber(letter: string): number {
  let result = 0;
  for (let i = 0; i < letter.length; i++) {
    result = result * 26 + letter.charCodeAt(i) - 64;
  }
  return result;
}

/**
 * Convert column number (1, 2, 3) to column letter (A, B, C)
 */
export function columnNumberToLetter(num: number): string {
  let letter = '';
  while (num > 0) {
    num--;
    letter = String.fromCharCode(65 + (num % 26)) + letter;
    num = Math.floor(num / 26);
  }
  return letter;
}

/**
 * Extract cell value exactly as it appears in Excel
 */
function cloneCellValue(value: any): any {
  if (value === null || value === undefined) return value;

  if (value instanceof Date) {
    return new Date(value.getTime());
  }

  if (typeof value !== 'object') {
    return value;
  }

  // ExcelJS uses plain objects for some cell value types (e.g. hyperlink, richText)
  // Shallow clone is typically enough to avoid sharing references.
  if (Array.isArray(value)) {
    return value.map(v => cloneCellValue(v));
  }

  if ('richText' in (value as any) && Array.isArray((value as any).richText)) {
    return {
      ...(value as any),
      richText: (value as any).richText.map((segment: any) => ({ ...segment }))
    };
  }

  return { ...(value as any) };
}

export function extractSafeCellValue(cell: ExcelJS.Cell): any {
  const value = cell.value;
  if (value === null || value === undefined) return null;

  if (cell.type === ExcelJS.ValueType.Formula) {
    return cell.result ?? null;
  }

  if (cell.type === ExcelJS.ValueType.RichText) {
    const richText = value as ExcelJS.CellRichTextValue;
    return richText.richText.map(segment => segment.text).join('');
  }

  if (typeof value === 'string') {
    // Do not normalize/parse date-like strings.
    // Preserve original text exactly (e.g. "1/29/2026", "2026-01-29").
    return value;
  }

  // Keep numeric value as-is - this handles Excel date serial numbers
  // The numFmt (number format) will be copied separately to preserve display format
  if (typeof value === 'number') {
    return value;
  }

  if (value instanceof Date) {
    return value;
  }

  if (typeof value === 'object') {
    if ('text' in (value as any)) {
      return (value as any).text;
    }
    if ('result' in (value as any)) {
      return (value as any).result ?? null;
    }
  }

  return cloneCellValue(value);
}

/**
 * Copy row with formatting - keeps original column positions
 */
export function copyRowWithFormatting(sourceRow: ExcelJS.Row, targetRow: ExcelJS.Row, startCol: number, endCol: number): void {
  // Copy cells keeping original column positions
  for (let colIndex = startCol; colIndex <= endCol; colIndex++) {
    const sourceCell = sourceRow.getCell(colIndex);
    const targetCell = targetRow.getCell(colIndex); // Keep same column position!

    // Extract safe cell value using helper function
    const cellValue = extractSafeCellValue(sourceCell);
    if (cellValue !== null) {
      targetCell.value = cellValue;
    }

    // Copy formatting with deep cloning to prevent cross-workbook reference issues
    if (sourceCell.font) {
      targetCell.font = cloneStyle(sourceCell.font);
    }
    if (sourceCell.fill) {
      targetCell.fill = cloneStyle(sourceCell.fill);
    }
    if (sourceCell.alignment) {
      targetCell.alignment = cloneStyle(sourceCell.alignment);
    }
    if (sourceCell.border) {
      targetCell.border = cloneStyle(sourceCell.border);
    }
    // Copy number format if it exists
    if (sourceCell.numFmt) {
      targetCell.numFmt = sourceCell.numFmt;
    }
  }
}

/**
 * Copy column widths from source sheet to target sheet
 */
export function copyColumnWidths(sourceSheet: ExcelJS.Worksheet, targetSheet: ExcelJS.Worksheet, startCol: number, endCol: number): void {
  for (let colIndex = startCol; colIndex <= endCol; colIndex++) {
    const sourceCol = sourceSheet.getColumn(colIndex);
    const targetCol = targetSheet.getColumn(colIndex);
    if (sourceCol.width) {
      targetCol.width = sourceCol.width;
    }
  }
}

function copyWorksheetContents(sourceSheet: ExcelJS.Worksheet, targetSheet: ExcelJS.Worksheet): void {
  const sourceColumnCount = sourceSheet.columnCount || 0;
  if (sourceColumnCount > 0) {
    copyColumnWidths(sourceSheet, targetSheet, 1, sourceColumnCount);
  }

  sourceSheet.eachRow({ includeEmpty: true }, row => {
    const targetRow = targetSheet.getRow(row.number);
    if (row.height) {
      targetRow.height = row.height;
    }

    row.eachCell({ includeEmpty: true }, (cell, colNum) => {
      const newCell = targetRow.getCell(colNum);

      const safeValue = extractSafeCellValue(cell);
      if (safeValue !== null) {
        newCell.value = safeValue;
      }

      if (cell.font) newCell.font = cell.font;
      if (cell.fill) newCell.fill = cell.fill;
      if (cell.alignment) newCell.alignment = cell.alignment;
      if (cell.border) newCell.border = cell.border;
      if (cell.numFmt) newCell.numFmt = cell.numFmt;
    });
  });

  const merges = sourceSheet.model.merges || [];
  for (const merge of merges) {
    try {
      targetSheet.mergeCells(merge);
    } catch {
      // ignore
    }
  }
}

/**
 * Copy merged cells from source sheet to target sheet for specific rows
 */
export function copyMergedCells(sourceSheet: ExcelJS.Worksheet, targetSheet: ExcelJS.Worksheet, sourceRowStart: number, sourceRowEnd: number, targetRowOffset: number, startCol: number, endCol: number): void {
  // Get all merged cell ranges from source
  const merges = sourceSheet.model.merges || [];

  for (const merge of merges) {
    // Parse merge range like "A1:B2"
    const match = merge.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/);
    if (!match) continue;

    const startColLetter = match[1];
    const startRowNum = parseInt(match[2]);
    const endColLetter = match[3];
    const endRowNum = parseInt(match[4]);

    const mergeStartCol = columnLetterToNumber(startColLetter);
    const mergeEndCol = columnLetterToNumber(endColLetter);

    // Check if this merge is within our row and column range
    if (startRowNum >= sourceRowStart && endRowNum <= sourceRowEnd && mergeStartCol >= startCol && mergeEndCol <= endCol) {
      // Calculate new row positions
      const newStartRow = startRowNum - sourceRowStart + targetRowOffset;
      const newEndRow = endRowNum - sourceRowStart + targetRowOffset;

      try {
        targetSheet.mergeCells(newStartRow, mergeStartCol, newEndRow, mergeEndCol);
      } catch (e) {
        // Ignore merge errors (cell might already be merged)
      }
    }
  }
}

/**
 * Deep clone style objects to prevent corruption when merging across workbooks
 */
function cloneStyle(style: any): any {
  if (!style) return undefined;
  try {
    return JSON.parse(JSON.stringify(style));
  } catch (e) {
    return { ...style };
  }
}

/**
 * Check if a file is a CSV file based on its name or MIME type
 */
export function isCsvFile(file: any): boolean {
  if (!file) return false;
  const name = (typeof file === 'string' ? file : file.name || '').trim().toLowerCase();
  const type = (typeof file === 'object' && file.type ? file.type : '').toLowerCase();
  return name.endsWith('.csv') || type === 'text/csv' || type === 'application/csv' || type.includes('csv');
}

/**
 * Determine the output format based on user preference and input files.
 * If user preference is 'csv' or 'xlsx', respect it.
 * If 'auto' or undefined:
 * - If any input file is CSV, output as 'csv'.
 * - Otherwise (only Excel files), output as 'xlsx'.
 */
export function resolveOutputFormat(
  files: File[],
  formatPreference?: 'auto' | 'xlsx' | 'csv'
): 'xlsx' | 'csv' {
  if (formatPreference === 'csv') return 'csv';
  if (formatPreference === 'xlsx') return 'xlsx';
  if (files && files.length > 0 && files.some(isCsvFile)) {
    return 'csv';
  }
  return 'xlsx';
}

function formatCsvCell(val: any): string {
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val.text) return formatCsvCell(val.text);
    if (val.result !== undefined) return formatCsvCell(val.result);
    if (val instanceof Date) return val.toISOString();
  }
  let str = String(val);
  if (str.includes('"') || str.includes(',') || str.includes('\n') || str.includes('\r')) {
    return `"${str.replace(/"/g, '""')}"`;
  }
  return str;
}

/**
 * Export an ExcelJS worksheet to a standard RFC 4180 CSV byte array.
 * Works natively in both browser and Node.js without stream dependencies.
 */
export function worksheetToCsv(worksheet: ExcelJS.Worksheet): Uint8Array {
  const lines: string[] = [];
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    const cells: string[] = [];
    const count = row.cellCount;
    for (let c = 1; c <= count; c++) {
      const cell = row.getCell(c);
      cells.push(formatCsvCell(extractSafeCellValue(cell)));
    }
    lines.push(cells.join(','));
  });
  return new TextEncoder().encode(lines.join('\r\n'));
}

/**
 * Prepend UTF-8 BOM (0xEF, 0xBB, 0xBF) to ensure Excel and other tools
 * correctly display Vietnamese/Unicode characters when opening CSV files.
 */
export function addUtf8Bom(data: Uint8Array | Buffer | ArrayBuffer): Uint8Array {
  const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return bytes;
  }
  const withBom = new Uint8Array(bytes.length + 3);
  withBom[0] = 0xef;
  withBom[1] = 0xbb;
  withBom[2] = 0xbf;
  withBom.set(bytes, 3);
  return withBom;
}

function castCsvValue(val: string): any {
  const trimmed = val.trim();
  if (trimmed === '') return '';
  // Check if it's a number
  if (/^-?\d+(\.\d+)?$/.test(trimmed)) {
    // Preserve leading zeroes like 0123 or 0901234567 as string
    if (trimmed.length > 1 && trimmed.startsWith('0') && !trimmed.startsWith('0.')) {
      return val;
    }
    if (trimmed.length > 2 && trimmed.startsWith('-0') && !trimmed.startsWith('-0.')) {
      return val;
    }
    const num = Number(trimmed);
    if (!isNaN(num)) return num;
  }
  return val;
}

/**
 * Parse CSV text into 2D array adhering to RFC 4180
 * Handles commas, semicolons, tabs, quotes, escaped quotes, multiline text, and UTF-8 BOM
 */
export function parseCSV(text: string): any[][] {
  if (!text) return [];

  // Strip UTF-8 BOM if present
  if (text.charCodeAt(0) === 0xFEFF) {
    text = text.slice(1);
  }

  // Detect delimiter: comma, semicolon, or tab
  let delimiter = ',';
  const firstLine = text.split(/\r\n|\n|\r/)[0] || '';
  const commaCount = (firstLine.match(/,/g) || []).length;
  const semiCount = (firstLine.match(/;/g) || []).length;
  const tabCount = (firstLine.match(/\t/g) || []).length;
  if (semiCount > commaCount && semiCount > tabCount) {
    delimiter = ';';
  } else if (tabCount > commaCount && tabCount > semiCount) {
    delimiter = '\t';
  }

  const rows: any[][] = [];
  let currentRow: any[] = [];
  let currentVal = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (inQuotes) {
      if (char === '"' && nextChar === '"') {
        currentVal += '"';
        i++; // skip escaped quote
      } else if (char === '"') {
        inQuotes = false;
      } else {
        currentVal += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === delimiter) {
        currentRow.push(castCsvValue(currentVal));
        currentVal = '';
      } else if (char === '\r') {
        if (nextChar === '\n') {
          i++;
        }
        currentRow.push(castCsvValue(currentVal));
        rows.push(currentRow);
        currentRow = [];
        currentVal = '';
      } else if (char === '\n') {
        currentRow.push(castCsvValue(currentVal));
        rows.push(currentRow);
        currentRow = [];
        currentVal = '';
      } else {
        currentVal += char;
      }
    }
  }

  if (currentVal || currentRow.length > 0) {
    currentRow.push(castCsvValue(currentVal));
    rows.push(currentRow);
  }

  return rows;
}

/**
 * Load Excel or CSV file into an ExcelJS Workbook
 */
export async function loadWorkbookFromFile(file: File): Promise<ExcelJS.Workbook> {
  const workbook = new ExcelJS.Workbook();
  if (isCsvFile(file)) {
    const text = typeof file.text === 'function'
      ? await file.text()
      : new TextDecoder('utf-8').decode(await file.arrayBuffer());
    const rows = parseCSV(text);
    const rawName = file.name.replace(/\.csv$/i, '').trim();
    const safeSheetName = rawName.replace(/[\\/?*:[\]]/g, '_').slice(0, 31) || 'Sheet1';
    const worksheet = workbook.addWorksheet(safeSheetName);
    for (const row of rows) {
      worksheet.addRow(row);
    }
  } else {
    const arrayBuffer = await file.arrayBuffer();
    await workbook.xlsx.load(new Uint8Array(arrayBuffer) as any);
  }
  return workbook;
}

/**
 * Merge multiple Excel files into one
 */
export async function mergeExcelFiles(files: File[], config: MergeConfig): Promise<Buffer> {
  const outputWorkbook = new ExcelJS.Workbook();
  const consolidatedSheet = outputWorkbook.addWorksheet('Consolidated');

  const startColNum = columnLetterToNumber(config.startColumn);
  const endColNum = columnLetterToNumber(config.endColumn);
  const startRowNum = config.startRow;

  let targetRowNum = 1;
  let firstSheetOfAll = true;
  let signatureRowsToAdd: { row: ExcelJS.Row; sourceSheet: ExcelJS.Worksheet }[] = [];
  let usedSheetNames = new Set<string>(['Consolidated']);

  for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
    const file = files[fileIndex];
    const sourceWorkbook = await loadWorkbookFromFile(file);

    for (const sourceSheet of sourceWorkbook.worksheets) {
      if (sourceSheet.name === 'Consolidated') continue;

      // 1. Preserve the original sheet in the output workbook
      let originalSheetName = sourceSheet.name;
      let uniqueName = originalSheetName;
      let counter = 1;
      while (usedSheetNames.has(uniqueName)) {
        uniqueName = `${originalSheetName} (${counter})`;
        counter++;
      }
      usedSheetNames.add(uniqueName);

      const newSheet = outputWorkbook.addWorksheet(uniqueName);
      copyWorksheetContents(sourceSheet, newSheet);

      // 2. Process data for the Consolidated sheet
      if (firstSheetOfAll) {
        copyColumnWidths(sourceSheet, consolidatedSheet, startColNum, endColNum);
      }

      const totalRowNum = findTotalRow(sourceSheet);
      const signatureRowNum = findSignatureRow(sourceSheet);

      // Copy header rows only from the very first sheet
      if (firstSheetOfAll) {
        for (let rowNum = 1; rowNum < startRowNum; rowNum++) {
          const sourceRow = sourceSheet.getRow(rowNum);
          const targetRow = consolidatedSheet.getRow(targetRowNum);
          if (sourceRow.height) targetRow.height = sourceRow.height;

          copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
          targetRowNum++;
        }
        copyMergedCells(sourceSheet, consolidatedSheet, 1, startRowNum - 1, 1, startColNum, endColNum);
      }

      // Copy data rows to consolidated sheet
      const dataStartRow = firstSheetOfAll ? startRowNum : startRowNum + 1;
      const lastRow = sourceSheet.lastRow?.number || 0;

      for (let rowNum = dataStartRow; rowNum <= lastRow; rowNum++) {
        const sourceRow = sourceSheet.getRow(rowNum);

        // Handle total row
        if (totalRowNum && rowNum === totalRowNum) {
          if (config.includeTotal) {
            const targetRow = consolidatedSheet.getRow(targetRowNum);
            if (sourceRow.height) targetRow.height = sourceRow.height;
            copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
            targetRowNum++;
          }
          continue;
        }

        // Handle signature section
        if (signatureRowNum && rowNum >= signatureRowNum) {
          if (config.includeSignature && signatureRowsToAdd.length === 0) {
            for (let sigRowNum = signatureRowNum; sigRowNum <= lastRow; sigRowNum++) {
              signatureRowsToAdd.push({
                row: sourceSheet.getRow(sigRowNum),
                sourceSheet
              });
            }
          }
          break;
        }

        // Skip subtotal rows
        if (!config.includeTotal && isSubtotalRow(sourceRow, startColNum, endColNum)) {
          continue;
        }

        // Copy regular data row
        const targetRow = consolidatedSheet.getRow(targetRowNum);
        if (sourceRow.height) targetRow.height = sourceRow.height;
        copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
        targetRowNum++;
      }

      if (firstSheetOfAll) firstSheetOfAll = false;
    }
  }

  // Add signature section at the end of Consolidated sheet
  if (signatureRowsToAdd.length > 0) {
    for (const { row: sigRow } of signatureRowsToAdd) {
      const targetRow = consolidatedSheet.getRow(targetRowNum);
      if (sigRow.height) targetRow.height = sigRow.height;
      copyRowWithFormatting(sigRow, targetRow, startColNum, endColNum);
      targetRowNum++;
    }
  }

  const outputFormat = resolveOutputFormat(files, config.outputFormat);
  if (outputFormat === 'csv') {
    return worksheetToCsv(consolidatedSheet) as any;
  }

  return await outputWorkbook.xlsx.writeBuffer() as any;
}

/**
 * Read sheet names from an Excel or CSV file (for preview purposes)
 */
export async function readSheetNames(file: File): Promise<string[]> {
  if (isCsvFile(file)) {
    const rawName = file.name.replace(/\.csv$/i, '').trim();
    const safeSheetName = rawName.replace(/[\\/?*:[\]]/g, '_').slice(0, 31) || 'Sheet1';
    return [safeSheetName];
  }
  const arrayBuffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(new Uint8Array(arrayBuffer) as any);
  return workbook.worksheets.map(ws => ws.name);
}

/**
 * Merge multiple Excel files into one file containing:
 * 1. All original sheets from every file (renamed if duplicates)
 * 2. A "Tổng hợp" (Consolidated) sheet that merges data from all sheets
 *    using the same logic as mergeExcelFiles()
 */
export async function mergeMultipleFiles(files: File[], config: MergeConfig): Promise<Buffer> {
  const outputWorkbook = new ExcelJS.Workbook();
  const consolidatedSheet = outputWorkbook.addWorksheet('Tổng hợp');

  const startColNum = columnLetterToNumber(config.startColumn);
  const endColNum = columnLetterToNumber(config.endColumn);
  const startRowNum = config.startRow;

  let targetRowNum = 1;
  let firstSheetOfAll = true;
  let signatureRowsToAdd: { row: ExcelJS.Row; sourceSheet: ExcelJS.Worksheet }[] = [];
  let usedSheetNames = new Set<string>(['Tổng hợp']);

  for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
    const file = files[fileIndex];
    const sourceWorkbook = await loadWorkbookFromFile(file);

    // Get base name of the file (without extension) for naming context
    const fileBaseName = file.name.replace(/\.(xlsx|xls|csv)$/i, '');

    for (const sourceSheet of sourceWorkbook.worksheets) {
      // 1. Preserve the original sheet in the output workbook
      let originalSheetName = sourceSheet.name;
      let uniqueName = originalSheetName;
      let counter = 1;
      while (usedSheetNames.has(uniqueName)) {
        uniqueName = `${originalSheetName} (${counter})`;
        counter++;
      }
      usedSheetNames.add(uniqueName);

      const newSheet = outputWorkbook.addWorksheet(uniqueName);
      copyWorksheetContents(sourceSheet, newSheet);

      // 2. Process data for the Consolidated (Tổng hợp) sheet
      if (firstSheetOfAll) {
        copyColumnWidths(sourceSheet, consolidatedSheet, startColNum, endColNum);
      }

      const totalRowNum = findTotalRow(sourceSheet);
      const signatureRowNum = findSignatureRow(sourceSheet);

      // Copy header rows only from the very first sheet
      if (firstSheetOfAll) {
        for (let rowNum = 1; rowNum < startRowNum; rowNum++) {
          const sourceRow = sourceSheet.getRow(rowNum);
          const targetRow = consolidatedSheet.getRow(targetRowNum);
          if (sourceRow.height) targetRow.height = sourceRow.height;

          copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
          targetRowNum++;
        }
        copyMergedCells(sourceSheet, consolidatedSheet, 1, startRowNum - 1, 1, startColNum, endColNum);
      }

      // Copy data rows to consolidated sheet
      const dataStartRow = firstSheetOfAll ? startRowNum : startRowNum + 1;
      const lastRow = sourceSheet.lastRow?.number || 0;

      for (let rowNum = dataStartRow; rowNum <= lastRow; rowNum++) {
        const sourceRow = sourceSheet.getRow(rowNum);

        // Handle total row
        if (totalRowNum && rowNum === totalRowNum) {
          if (config.includeTotal) {
            const targetRow = consolidatedSheet.getRow(targetRowNum);
            if (sourceRow.height) targetRow.height = sourceRow.height;
            copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
            targetRowNum++;
          }
          continue;
        }

        // Handle signature section
        if (signatureRowNum && rowNum >= signatureRowNum) {
          if (config.includeSignature && signatureRowsToAdd.length === 0) {
            for (let sigRowNum = signatureRowNum; sigRowNum <= lastRow; sigRowNum++) {
              signatureRowsToAdd.push({
                row: sourceSheet.getRow(sigRowNum),
                sourceSheet
              });
            }
          }
          break;
        }

        // Skip subtotal rows
        if (!config.includeTotal && isSubtotalRow(sourceRow, startColNum, endColNum)) {
          continue;
        }

        // Copy regular data row
        const targetRow = consolidatedSheet.getRow(targetRowNum);
        if (sourceRow.height) targetRow.height = sourceRow.height;
        copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
        targetRowNum++;
      }

      if (firstSheetOfAll) firstSheetOfAll = false;
    }
  }

  // Add signature section at the end of Consolidated sheet
  if (signatureRowsToAdd.length > 0) {
    for (const { row: sigRow } of signatureRowsToAdd) {
      const targetRow = consolidatedSheet.getRow(targetRowNum);
      if (sigRow.height) targetRow.height = sigRow.height;
      copyRowWithFormatting(sigRow, targetRow, startColNum, endColNum);
      targetRowNum++;
    }
  }

  // Move Tổng hợp sheet to first position
  // ExcelJS doesn't have a native moveSheet, so we reorder by setting orderNo
  // (orderNo exists at runtime but is not in the public type definitions)
  const worksheets = outputWorkbook.worksheets;
  const consolidatedIndex = worksheets.findIndex(ws => ws.name === 'Tổng hợp');
  if (consolidatedIndex > 0) {
    // Reorder: set orderNo so Tổng hợp comes first
    (consolidatedSheet as any).orderNo = 0;
    let order = 1;
    for (const ws of worksheets) {
      if (ws.name !== 'Tổng hợp') {
        (ws as any).orderNo = order++;
      }
    }
  }

  const outputFormat = resolveOutputFormat(files, config.outputFormat);
  if (outputFormat === 'csv') {
    return worksheetToCsv(consolidatedSheet) as any;
  }

  return await outputWorkbook.xlsx.writeBuffer() as any;
}
