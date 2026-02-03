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
 * Find signature row - looks for signature keywords
 * Supports: "Người lập phiếu", "Người nhận", "Signature of sender", "Chữ ký người gửi"
 */
export function findSignatureRow(worksheet: ExcelJS.Worksheet): number | null {
  let signatureRow: number | null = null;

  worksheet.eachRow((row, rowNumber) => {
    if (signatureRow) return; // Already found

    row.eachCell((cell) => {
      const cellValue = cell.value?.toString().trim().toLowerCase();
      if (!cellValue) return;

      // Check for signature keywords
      const signatureKeywords = [
        'người lập phiếu',
        'người nhận',
        'signature of sender',
        'chữ ký người gửi',
        'signature',
        'chữ ký'
      ];

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
 * Copy row with formatting - keeps original column positions
 */
export function copyRowWithFormatting(
  sourceRow: ExcelJS.Row,
  targetRow: ExcelJS.Row,
  startCol: number,
  endCol: number
): void {
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
    if (sourceCell.numFmt) {
      targetCell.numFmt = sourceCell.numFmt;
    }
  }
}

/**
 * Copy column widths from source sheet to target sheet
 */
export function copyColumnWidths(
  sourceSheet: ExcelJS.Worksheet,
  targetSheet: ExcelJS.Worksheet,
  startCol: number,
  endCol: number
): void {
  for (let colIndex = startCol; colIndex <= endCol; colIndex++) {
    const sourceCol = sourceSheet.getColumn(colIndex);
    const targetCol = targetSheet.getColumn(colIndex);
    if (sourceCol.width) {
      targetCol.width = sourceCol.width;
    }
  }
}

/**
 * Copy merged cells from source sheet to target sheet for specific rows
 */
export function copyMergedCells(
  sourceSheet: ExcelJS.Worksheet,
  targetSheet: ExcelJS.Worksheet,
  sourceRowStart: number,
  sourceRowEnd: number,
  targetRowOffset: number,
  startCol: number,
  endCol: number
): void {
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
    if (startRowNum >= sourceRowStart && endRowNum <= sourceRowEnd &&
        mergeStartCol >= startCol && mergeEndCol <= endCol) {
      // Calculate new row positions
      const newStartRow = startRowNum - sourceRowStart + targetRowOffset;
      const newEndRow = endRowNum - sourceRowStart + targetRowOffset;
      
      try {
        targetSheet.mergeCells(
          newStartRow, mergeStartCol,
          newEndRow, mergeEndCol
        );
      } catch (e) {
        // Ignore merge errors (cell might already be merged)
      }
    }
  }
}

/**
 * Merge multiple Excel files into one
 */
export async function mergeExcelFiles(
  files: File[],
  config: MergeConfig
): Promise<ArrayBuffer> {
  // Load the first file as base workbook
  const firstFileBuffer = await files[0].arrayBuffer();
  const baseWorkbook = new ExcelJS.Workbook();
  await baseWorkbook.xlsx.load(firstFileBuffer);

  // Get the first sheet as template for structure
  const templateSheet = baseWorkbook.worksheets[0];
  
  // Create a new consolidated worksheet
  const worksheet = baseWorkbook.addWorksheet('Consolidated');

  const startColNum = columnLetterToNumber(config.startColumn);
  const endColNum = columnLetterToNumber(config.endColumn);
  const startRowNum = config.startRow;

  // Copy column widths from template sheet
  copyColumnWidths(templateSheet, worksheet, startColNum, endColNum);

  let targetRowNum = 1;
  let isFirstSheet = true;
  let signatureRowsToAdd: { row: ExcelJS.Row; sourceSheet: ExcelJS.Worksheet }[] = [];
  let firstSourceSheet: ExcelJS.Worksheet | null = null;

  for (const file of files) {
    const arrayBuffer = await file.arrayBuffer();
    const sourceWorkbook = new ExcelJS.Workbook();
    await sourceWorkbook.xlsx.load(arrayBuffer);

    // Process each sheet in the file
    sourceWorkbook.worksheets.forEach((sourceSheet) => {
      // Save first source sheet for merged cells reference
      if (!firstSourceSheet) {
        firstSourceSheet = sourceSheet;
      }

      // Always find special rows so we can skip or include them based on config
      const totalRowNum = findTotalRow(sourceSheet);
      const signatureRowNum = findSignatureRow(sourceSheet);

      // For first sheet: copy header rows (1 to startRowNum-1) first
      if (isFirstSheet) {
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

      // Copy data rows (from startRowNum onwards)
      const lastRow = sourceSheet.lastRow?.number || 0;
      for (let rowNum = startRowNum; rowNum <= lastRow; rowNum++) {
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

      // After processing first sheet, mark it as done
      isFirstSheet = false;
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
