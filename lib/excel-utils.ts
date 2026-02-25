import ExcelJS from 'exceljs';

export interface MergeConfig {
  includeTotal: boolean;
  startRow: number;
  startColumn: string;
  endColumn: string;
  includeSignature: boolean;
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

    // Copy formatting
    if (sourceCell.font) {
      targetCell.font = { ...sourceCell.font };
    }
    if (sourceCell.fill) {
      targetCell.fill = { ...sourceCell.fill };
    }
    if (sourceCell.alignment) {
      targetCell.alignment = { ...sourceCell.alignment };
    }
    if (sourceCell.border) {
      targetCell.border = { ...sourceCell.border };
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

      if (cell.font) newCell.font = { ...cell.font };
      if (cell.fill) newCell.fill = { ...cell.fill };
      if (cell.alignment) newCell.alignment = { ...cell.alignment };
      if (cell.border) newCell.border = { ...cell.border };
      if (cell.numFmt) newCell.numFmt = cell.numFmt;
    });

    targetRow.commit();
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
 * Merge multiple Excel files into one
 */
export async function mergeExcelFiles(files: File[], config: MergeConfig): Promise<ArrayBuffer> {
  // Load the first file to get the template structure
  const firstFileBuffer = await files[0].arrayBuffer();
  const templateWorkbook = new ExcelJS.Workbook();
  await templateWorkbook.xlsx.load(firstFileBuffer, {
    ignoreNodes: ['sharedFormula']
  } as any);

  // Keep track of the first worksheet for formatting metadata
  const templateSheet = templateWorkbook.worksheets[0];

  // Add/replace consolidated sheet inside the first imported workbook
  const existingConsolidated = templateWorkbook.getWorksheet('Consolidated');
  if (existingConsolidated) {
    templateWorkbook.removeWorksheet(existingConsolidated.id);
  }
  const worksheet = templateWorkbook.addWorksheet('Consolidated');

  const startColNum = columnLetterToNumber(config.startColumn);
  const endColNum = columnLetterToNumber(config.endColumn);
  const startRowNum = config.startRow;

  // Copy column widths from template sheet
  copyColumnWidths(templateSheet, worksheet, startColNum, endColNum);

  let targetRowNum = 1;
  let firstSheetOfAll = true;
  let signatureRowsToAdd: { row: ExcelJS.Row; sourceSheet: ExcelJS.Worksheet }[] = [];

  for (let fileIndex = 0; fileIndex < files.length; fileIndex++) {
    const file = files[fileIndex];
    let sourceWorkbook: ExcelJS.Workbook;
    if (fileIndex === 0) {
      sourceWorkbook = templateWorkbook;
    } else {
      const arrayBuffer = await file.arrayBuffer();
      sourceWorkbook = new ExcelJS.Workbook();
      await sourceWorkbook.xlsx.load(arrayBuffer, {
        ignoreNodes: ['sharedFormula']
      } as any);
    }

    // Process all sheets in each file
    sourceWorkbook.worksheets.forEach(sourceSheet => {
      // Skip the "Consolidated" sheet itself to prevent duplication
      if (sourceSheet.name === 'Consolidated') return;
      // Always find special rows so we can skip or include them based on config
      const totalRowNum = findTotalRow(sourceSheet);
      const signatureRowNum = findSignatureRow(sourceSheet);

      // For first sheet of all: copy header rows (1 to startRowNum-1) first
      if (firstSheetOfAll) {
        for (let rowNum = 1; rowNum < startRowNum; rowNum++) {
          const sourceRow = sourceSheet.getRow(rowNum);
          const targetRow = worksheet.getRow(targetRowNum);
          copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
          targetRow.commit();
          targetRowNum++;
        }

        // Copy merged cells for header section
        copyMergedCells(sourceSheet, worksheet, 1, startRowNum - 1, 1, startColNum, endColNum);
      }

      // Copy data rows
      // For first sheet of all: copy from startRowNum (includes header row of data table)
      // For other sheets: copy from startRowNum + 1 (skip header row of data table)
      const dataStartRow = firstSheetOfAll ? startRowNum : startRowNum + 1;
      const lastRow = sourceSheet.lastRow?.number || 0;
      for (let rowNum = dataStartRow; rowNum <= lastRow; rowNum++) {
        const sourceRow = sourceSheet.getRow(rowNum);

        // Handle total row - skip if not included, add if included
        if (totalRowNum && rowNum === totalRowNum) {
          if (config.includeTotal) {
            const targetRow = worksheet.getRow(targetRowNum);
            copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
            targetRow.commit();
            targetRowNum++;
          }
          continue;
        }

        // Handle signature section - save to add only once at end
        if (signatureRowNum && rowNum >= signatureRowNum) {
          if (config.includeSignature && signatureRowsToAdd.length === 0) {
            // Save all signature rows from first sheet that has them
            for (let sigRowNum = signatureRowNum; sigRowNum <= lastRow; sigRowNum++) {
              signatureRowsToAdd.push({
                row: sourceSheet.getRow(sigRowNum),
                sourceSheet
              });
            }
          }
          break; // Stop processing this sheet's rows
        }

        // Skip subtotal/group total rows if includeTotal is false
        if (!config.includeTotal && isSubtotalRow(sourceRow, startColNum, endColNum)) {
          continue;
        }

        // Copy regular data row
        const targetRow = worksheet.getRow(targetRowNum);
        copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
        targetRow.commit();
        targetRowNum++;
      }

      // After processing first sheet of all, mark it as done
      if (firstSheetOfAll) {
        firstSheetOfAll = false;
      }
    });

  }

  // Add signature section only once at the very end
  if (signatureRowsToAdd.length > 0) {
    const sigStartRow = targetRowNum;
    for (const { row: sigRow } of signatureRowsToAdd) {
      const targetRow = worksheet.getRow(targetRowNum);
      copyRowWithFormatting(sigRow, targetRow, startColNum, endColNum);
      targetRow.commit();
      targetRowNum++;
    }
  }

  // Return the workbook with original sheets and the new consolidated one
  const buffer = await templateWorkbook.xlsx.writeBuffer();
  return buffer as unknown as ArrayBuffer;
}
