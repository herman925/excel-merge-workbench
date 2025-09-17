import * as XLSX from 'xlsx';

/**
 * Creates a map of all cells covered by merged ranges pointing to their top-left cell address
 */
export function createMergeMap(worksheet: XLSX.WorkSheet): Map<string, string> {
  const mergeMap = new Map<string, string>();
  
  if (!worksheet['!merges']) {
    return mergeMap;
  }

  worksheet['!merges'].forEach(merge => {
    const topLeftAddress = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
    
    // Map all cells in the merge range to the top-left cell
    for (let row = merge.s.r; row <= merge.e.r; row++) {
      for (let col = merge.s.c; col <= merge.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        mergeMap.set(cellAddress, topLeftAddress);
      }
    }
  });

  return mergeMap;
}

/**
 * Gets the value from a cell, handling merged cells by checking the top-left cell of the merge
 */
export function getCellValue(worksheet: XLSX.WorkSheet, cellAddress: string, mergeMap: Map<string, string>): any {
  // First try to get the direct cell value
  const cell = worksheet[cellAddress];
  if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
    return cell.v;
  }

  // If no direct value, check if this cell is part of a merged range
  const mergedCellAddress = mergeMap.get(cellAddress);
  if (mergedCellAddress) {
    const mergedCell = worksheet[mergedCellAddress];
    if (mergedCell && mergedCell.v !== undefined && mergedCell.v !== null && mergedCell.v !== '') {
      return mergedCell.v;
    }
  }

  return null;
}

/**
 * Reads column headers from a worksheet, handling merged cells
 */
export function readColumnHeaders(worksheet: XLSX.WorkSheet, headerRow: number = 1): string[] {
  if (!worksheet['!ref']) {
    return [];
  }

  const range = XLSX.utils.decode_range(worksheet['!ref']);
  const mergeMap = createMergeMap(worksheet);
  const columns: string[] = [];
  
  const headerRowIndex = headerRow - 1; // Convert to 0-based index
  
  for (let col = range.s.c; col <= range.e.c; col++) {
    const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: col });
    const cellValue = getCellValue(worksheet, cellAddress, mergeMap);
    
    if (cellValue !== null && cellValue !== undefined && String(cellValue).trim() !== '') {
      columns.push(String(cellValue).trim());
    } else {
      // Fallback to generic column name
      columns.push(`Column ${String.fromCharCode(65 + col)}`);
    }
  }
  
  return columns;
}