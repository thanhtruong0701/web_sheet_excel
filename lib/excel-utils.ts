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
    row.eachCell((cell) => {
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
export function isSubtotalRow(
  row: ExcelJS.Row,
  startCol: number,
  endCol: number
): boolean {
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
 * Find signature row (contains both "Người lập phiếu" and "Người nhận")
 */
export function findSignatureRow(worksheet: ExcelJS.Worksheet): number | null {
  let signatureRow: number | null = null;

  worksheet.eachRow((row, rowNumber) => {
    let hasCreator = false;
    let hasReceiver = false;

    row.eachCell((cell) => {
      const cellValue = cell.value?.toString().trim();
      if (cellValue?.includes('Người lập phiếu')) hasCreator = true;
      if (cellValue?.includes('Người nhận')) hasReceiver = true;
    });

    if (hasCreator && hasReceiver && !signatureRow) {
      signatureRow = rowNumber;
    }
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
 * Extract a safe cell value - always returns primitive value, never formula objects
 */
export function extractSafeCellValue(cell: ExcelJS.Cell): any {
  if (cell.value === undefined || cell.value === null) {
    return null;
  }

  const cellType = cell.type;

  if (cellType === ExcelJS.ValueType.Formula) {
    // For formula cells, use the calculated result only
    return cell.result ?? null;
  }
  
  if (cellType === ExcelJS.ValueType.RichText) {
    // Handle rich text
    const richText = cell.value as ExcelJS.CellRichTextValue;
    if (richText.richText) {
      return richText.richText.map((rt: any) => rt.text).join('');
    }
    return null;
  }
  
  if (typeof cell.value === 'object' && cell.value !== null) {
    // Handle other complex objects
    const val = cell.value as any;
    if (val.result !== undefined) {
      return val.result;
    }
    if (val.text !== undefined) {
      return val.text;
    }
    if (val.richText) {
      return val.richText.map((rt: any) => rt.text).join('');
    }
    // Skip formula objects completely
    if (val.formula !== undefined || val.sharedFormula !== undefined) {
      return val.result ?? null;
    }
    return null;
  }
  
  // Simple values (string, number, boolean, date)
  return cell.value;
}

/**
 * Copy row with formatting
 */
export function copyRowWithFormatting(
  sourceRow: ExcelJS.Row,
  targetRow: ExcelJS.Row,
  startCol: number,
  endCol: number
): void {
  let targetColIndex = 1;

  for (let colIndex = startCol; colIndex <= endCol; colIndex++) {
    const sourceCell = sourceRow.getCell(colIndex);
    const targetCell = targetRow.getCell(targetColIndex);

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
    if (sourceCell.numFmt) {
      targetCell.numFmt = sourceCell.numFmt;
    }

    targetColIndex++;
  }
}

/**
 * Merge multiple Excel files into one
 */
export async function mergeExcelFiles(
  files: File[],
  config: MergeConfig
): Promise<ArrayBuffer> {
  // Load the first file as base workbook to add merged sheet into it
  const firstFileBuffer = await files[0].arrayBuffer();
  const baseWorkbook = new ExcelJS.Workbook();
  await baseWorkbook.xlsx.load(firstFileBuffer);

  // Create a new consolidated worksheet
  const worksheet = baseWorkbook.addWorksheet('Consolidated');

  const startColNum = columnLetterToNumber(config.startColumn);
  const endColNum = columnLetterToNumber(config.endColumn);

  let targetRowNum = 1;
  let isFirstSheet = true;
  let signatureRowsToAdd: ExcelJS.Row[] = []; // Store all signature section rows

  for (const file of files) {
    const arrayBuffer = await file.arrayBuffer();
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.load(arrayBuffer);

    // Process each sheet in the file
    sourceWorkbook.worksheets.forEach((sourceSheet) => {
      // Always find special rows so we can skip or include them based on config
      const totalRowNum = findTotalRow(sourceSheet);
      const signatureRowNum = findSignatureRow(sourceSheet);

      const startRowNum = config.startRow;

      // Copy rows from source sheet
      sourceSheet.eachRow((sourceRow, rowNum) => {
        // For first sheet: copy ALL rows (starting from row 1)
        // For subsequent sheets: skip header rows (rows before startRowNum), only copy data
        if (!isFirstSheet && rowNum <= startRowNum) {
          return; // Skip header rows for non-first sheets
        }

        // Handle total row - skip if not included, add if included
        if (totalRowNum && rowNum === totalRowNum) {
          if (config.includeTotal) {
            const targetRow = worksheet.addRow([]);
            copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
            targetRowNum++;
          }
          // Always skip TOTAL row from regular data copy (either included above or excluded)
          return;
        }

        // Handle signature section - skip always during data copy, save to add only once at end if included
        if (signatureRowNum && rowNum >= signatureRowNum) {
          // Save signature section rows only from the first sheet that has them
          if (config.includeSignature && signatureRowsToAdd.length === 0) {
            signatureRowsToAdd.push(sourceRow);
          } else if (config.includeSignature && signatureRowsToAdd.length > 0 && isFirstSheet) {
            // Continue adding rows from the same signature section
            signatureRowsToAdd.push(sourceRow);
          }
          // Skip all rows from signature row onwards (signature section)
          return;
        }

        // Skip subtotal/group total rows if includeTotal is false
        if (!config.includeTotal && isSubtotalRow(sourceRow, startColNum, endColNum)) {
          return;
        }

        const targetRow = worksheet.addRow([]);
        copyRowWithFormatting(sourceRow, targetRow, startColNum, endColNum);
        targetRowNum++;
      });

      // After processing first sheet, mark it as done
      isFirstSheet = false;
    });
  }

  // Add signature section only once at the very end
  if (signatureRowsToAdd.length > 0) {
    for (const sigRow of signatureRowsToAdd) {
      const targetRow = worksheet.addRow([]);
      copyRowWithFormatting(sigRow, targetRow, startColNum, endColNum);
    }
  }

  // Set column widths
  const colRange = endColNum - startColNum + 1;
  for (let i = 1; i <= colRange; i++) {
    worksheet.columns[i - 1].width = 15;
  }

  try {
    const buffer = await baseWorkbook.xlsx.writeBuffer();
    return buffer as unknown as ArrayBuffer;
  } catch (error: any) {
    console.log('[v0] Error in baseWorkbook, attempting clean rebuild:', error.message);
    
    // If there are still formula issues, completely rebuild without any formulas
    const cleanWorkbook = new ExcelJS.Workbook();
    
    // Copy all original sheets from first file
    const firstBuffer = await files[0].arrayBuffer();
    const firstWb = new ExcelJS.Workbook();
    await firstWb.xlsx.load(firstBuffer);
    
    firstWb.worksheets.forEach((sheet) => {
      const newSheet = cleanWorkbook.addWorksheet(sheet.name);
      sheet.eachRow((row) => {
        const newRow = newSheet.addRow([]);
        row.eachCell((cell, colNum) => {
          const newCell = newRow.getCell(colNum);
          
          // Extract safe value - never copy formula objects
          const safeValue = extractSafeCellValue(cell);
          if (safeValue !== null) {
            newCell.value = safeValue;
          }
          
          // Copy all formatting
          if (cell.font) newCell.font = { ...cell.font };
          if (cell.fill) newCell.fill = { ...cell.fill };
          if (cell.alignment) newCell.alignment = { ...cell.alignment };
          if (cell.border) newCell.border = { ...cell.border };
          if (cell.numFmt) newCell.numFmt = cell.numFmt;
        });
      });
    });

    // Add consolidated sheet from the problematic worksheet
    const cleanSheet = cleanWorkbook.addWorksheet('Consolidated');
    worksheet.eachRow((row) => {
      const newRow = cleanSheet.addRow([]);
      row.eachCell((cell, colNum) => {
        const newCell = newRow.getCell(colNum);
        
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
    });

    const buffer = await cleanWorkbook.xlsx.writeBuffer();
    return buffer as unknown as ArrayBuffer;
  }
}
