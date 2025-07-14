import * as XLSX from 'xlsx';
import { ExcelFile, WorksheetData, ColumnMapping } from '../components/ExcelCombiner';

export interface ProcessingResults {
  combinedData: any[][];
  duplicateRows: DuplicateRowInfo[];
  unmappedColumns: UnmappedColumnInfo[];
  totalRowsProcessed: number;
  duplicatesRemoved: number;
  successfulFiles: number;
  previewData: any[][];
}

export interface DuplicateRowInfo {
  originalRow: any[];
  duplicateCount: number;
  sourceFiles: string[];
  rowIndex: number;
}

export interface UnmappedColumnInfo {
  fileName: string;
  worksheetName: string;
  columnName: string;
  dataType: string;
  sampleValues: string[];
}

interface RowData {
  data: any[];
  sourceFile: string;
  sourceWorksheet: string;
  originalRowIndex: number;
}

export class ExcelProcessor {
  private files: ExcelFile[];
  private worksheets: WorksheetData[];
  private mappings: ColumnMapping[];
  private keyColumn?: string;

  constructor(files: ExcelFile[], worksheets: WorksheetData[], mappings: ColumnMapping[], keyColumn?: string) {
    this.files = files;
    this.worksheets = worksheets;
    this.mappings = mappings;
    this.keyColumn = keyColumn;
  }

  async processFiles(): Promise<ProcessingResults> {
    console.log('Starting Excel processing...');
    
    // Read all data from worksheets organized by file
    const fileDataMap = new Map<string, RowData[]>();
    const unmappedColumns: UnmappedColumnInfo[] = [];
    let successfulFiles = 0;

    for (const worksheet of this.worksheets) {
      const file = this.files.find(f => f.id === worksheet.fileId);
      if (!file) continue;

      try {
        const data = await this.readWorksheetData(file.file, worksheet);
        fileDataMap.set(worksheet.fileId, data);
        successfulFiles++;

        // Track unmapped columns
        const mappedColumns = this.mappings.flatMap(m => 
          m.mappings.filter(map => map.fileId === worksheet.fileId).map(map => map.column)
        );
        
        const unmapped = worksheet.columns.filter(col => !mappedColumns.includes(col));
        unmapped.forEach(col => {
          // Get sample values for this column
          const sampleValues = data.slice(0, 3).map(row => {
            const colIndex = worksheet.columns.indexOf(col);
            return row.data[colIndex] || '';
          }).filter(Boolean);

          unmappedColumns.push({
            fileName: file.name,
            worksheetName: worksheet.worksheetName,
            columnName: col,
            dataType: this.detectDataType(sampleValues),
            sampleValues: sampleValues.slice(0, 3)
          });
        });
      } catch (error) {
        console.error(`Failed to process ${file.name}:`, error);
      }
    }

    const totalRows = Array.from(fileDataMap.values()).reduce((sum, data) => sum + data.length, 0);
    console.log(`Processed ${totalRows} total rows from ${successfulFiles} files`);

    let combinedRowsData: RowData[];

    console.log('Global key column:', this.keyColumn);
    console.log('Individual worksheet key columns:', this.worksheets.map(w => ({ fileId: w.fileId, keyColumn: w.keyColumn })));

    // Check if any worksheet has a key column defined
    const hasAnyKeyColumn = this.keyColumn || this.worksheets.some(w => w.keyColumn);

    if (hasAnyKeyColumn) {
      // Key-based matching using individual worksheet key columns
      console.log('Using key-based matching');
      combinedRowsData = this.combineDataByWorksheetKeys(fileDataMap);
    } else {
      // Row position matching (fallback)
      console.log('Using row position matching (fallback)');
      combinedRowsData = this.combineDataByPosition(fileDataMap);
    }

    console.log('Combined rows count:', combinedRowsData.length);

    // Detect and handle duplicates
    const { uniqueData, duplicateInfo } = this.handleDuplicates(combinedRowsData);
    
    // Generate output headers
    const outputHeaders = this.mappings.map(m => m.outputColumn);
    const finalData = [outputHeaders, ...uniqueData];

    // Create preview (first 10 rows + headers)
    const previewData = finalData.slice(0, 11); // Headers + 10 data rows

    return {
      combinedData: finalData,
      duplicateRows: duplicateInfo,
      unmappedColumns,
      totalRowsProcessed: totalRows,
      duplicatesRemoved: duplicateInfo.length,
      successfulFiles,
      previewData
    };
  }

  private async readWorksheetData(file: File, worksheet: WorksheetData): Promise<RowData[]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const worksheetObj = workbook.Sheets[worksheet.worksheetName];
          
          if (!worksheetObj) {
            resolve([]);
            return;
          }

          // Convert to array of arrays, skipping header row
          const jsonData = XLSX.utils.sheet_to_json(worksheetObj, { 
            header: 1,
            defval: '',
            range: worksheet.headerRow // Skip header rows
          }) as any[][];

          // Filter out completely empty rows
          const filteredData = jsonData.filter(row => 
            row.some(cell => cell !== null && cell !== undefined && cell !== '')
          );

          const rowsData: RowData[] = filteredData.map((row, index) => ({
            data: row,
            sourceFile: this.files.find(f => f.id === worksheet.fileId)?.name || '',
            sourceWorksheet: worksheet.worksheetName,
            originalRowIndex: index + worksheet.headerRow + 1
          }));

          resolve(rowsData);
        } catch (error) {
          console.error('Error reading worksheet data:', error);
          reject(error);
        }
      };
      
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsArrayBuffer(file);
    });
  }

  private applyColumnMappings(allRowsData: RowData[]): RowData[] {
    return allRowsData.map(rowData => {
      const worksheet = this.worksheets.find(w => 
        this.files.find(f => f.id === w.fileId)?.name === rowData.sourceFile &&
        w.worksheetName === rowData.sourceWorksheet
      );

      if (!worksheet) return rowData;

      const mappedRowData: any[] = new Array(this.mappings.length).fill('');

      this.mappings.forEach((mapping, mappingIndex) => {
        const fileMapping = mapping.mappings.find(m => m.fileId === worksheet.fileId);
        if (fileMapping) {
          const columnIndex = worksheet.columns.indexOf(fileMapping.column);
          if (columnIndex >= 0 && columnIndex < rowData.data.length) {
            mappedRowData[mappingIndex] = rowData.data[columnIndex] || '';
          }
        }
      });

      return {
        ...rowData,
        data: mappedRowData
      };
    });
  }

  private handleDuplicates(mappedData: RowData[]): { uniqueData: any[][], duplicateInfo: DuplicateRowInfo[] } {
    const rowMap = new Map<string, RowData[]>();
    const duplicateInfo: DuplicateRowInfo[] = [];

    // Group rows by their data content
    mappedData.forEach(row => {
      const key = JSON.stringify(row.data);
      if (!rowMap.has(key)) {
        rowMap.set(key, []);
      }
      rowMap.get(key)!.push(row);
    });

    const uniqueData: any[][] = [];

    // Process each group
    rowMap.forEach((rows, key) => {
      if (rows.length > 1) {
        // This is a duplicate
        const sourceFiles = [...new Set(rows.map(r => r.sourceFile))];
        duplicateInfo.push({
          originalRow: rows[0].data,
          duplicateCount: rows.length - 1,
          sourceFiles,
          rowIndex: rows[0].originalRowIndex
        });
      }
      
      // Keep only one instance (the first one)
      uniqueData.push(rows[0].data);
    });

    return { uniqueData, duplicateInfo };
  }

  private combineDataByKey(fileDataMap: Map<string, RowData[]>): RowData[] {
    if (!this.keyColumn) {
      console.log('No key column specified, returning empty array');
      return [];
    }

    // Find the key column mapping
    const keyMapping = this.mappings.find(m => m.outputColumn === this.keyColumn);
    console.log('Looking for key mapping with output column:', this.keyColumn);
    console.log('Available output columns:', this.mappings.map(m => m.outputColumn));
    
    if (!keyMapping) {
      console.log('Key mapping not found for column:', this.keyColumn);
      console.log('Available mappings:', this.mappings.map(m => m.outputColumn));
      
      // Try to find a mapping where the key column is one of the mapped columns instead
      const alternativeMapping = this.mappings.find(m => 
        m.mappings.some(fm => fm.column === this.keyColumn)
      );
      
      if (alternativeMapping) {
        console.log('Found alternative mapping where key column is a source column:', alternativeMapping);
        // Use the alternative mapping but we need to handle this differently
        return this.combineDataByAlternativeKey(fileDataMap, this.keyColumn);
      }
      
      return [];
    }

    console.log('Found key mapping:', keyMapping);

    // Create a map of key values to combined rows
    const keyRowMap = new Map<string, any[]>();

    // First pass: collect all unique key values
    fileDataMap.forEach((fileRows, fileId) => {
      const worksheet = this.worksheets.find(w => w.fileId === fileId);
      if (!worksheet) return;

      const keyFileMapping = keyMapping.mappings.find(m => m.fileId === fileId);
      if (!keyFileMapping) return;

      const keyColumnIndex = worksheet.columns.indexOf(keyFileMapping.column);
      if (keyColumnIndex < 0) return;

      fileRows.forEach(row => {
        const keyValue = String(row.data[keyColumnIndex] || '').trim();
        if (keyValue && !keyRowMap.has(keyValue)) {
          keyRowMap.set(keyValue, new Array(this.mappings.length).fill(''));
        }
      });
    });

    // Second pass: populate the combined rows
    fileDataMap.forEach((fileRows, fileId) => {
      const worksheet = this.worksheets.find(w => w.fileId === fileId);
      if (!worksheet) return;

      const keyFileMapping = keyMapping.mappings.find(m => m.fileId === fileId);
      if (!keyFileMapping) return;

      const keyColumnIndex = worksheet.columns.indexOf(keyFileMapping.column);
      if (keyColumnIndex < 0) return;

      fileRows.forEach(row => {
        const keyValue = String(row.data[keyColumnIndex] || '').trim();
        if (!keyValue || !keyRowMap.has(keyValue)) return;

        const combinedRow = keyRowMap.get(keyValue)!;

        // Apply all column mappings for this file
        this.mappings.forEach((mapping, mappingIndex) => {
          const fileMapping = mapping.mappings.find(m => m.fileId === fileId);
          if (fileMapping) {
            const columnIndex = worksheet.columns.indexOf(fileMapping.column);
            if (columnIndex >= 0 && columnIndex < row.data.length) {
              const cellValue = row.data[columnIndex];
              if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
                combinedRow[mappingIndex] = cellValue;
              }
            }
          }
        });
      });
    });

    // Convert map to RowData array
    const combinedRowsData: RowData[] = [];
    keyRowMap.forEach((rowData, keyValue) => {
      if (rowData.some(cell => cell !== '')) {
        combinedRowsData.push({
          data: rowData,
          sourceFile: 'combined',
          sourceWorksheet: 'combined',
          originalRowIndex: combinedRowsData.length + 1
        });
      }
    });

    return combinedRowsData;
  }

  private combineDataByPosition(fileDataMap: Map<string, RowData[]>): RowData[] {
    // Row position matching (fallback)
    const maxRowCount = Math.max(...Array.from(fileDataMap.values()).map(data => data.length));
    const combinedRowsData: RowData[] = [];
    
    for (let rowIndex = 0; rowIndex < maxRowCount; rowIndex++) {
      const combinedRowData: any[] = new Array(this.mappings.length).fill('');
      
      // Apply column mappings for this row across all files
      this.mappings.forEach((mapping, mappingIndex) => {
        mapping.mappings.forEach(fileMapping => {
          const fileData = fileDataMap.get(fileMapping.fileId);
          if (fileData && fileData[rowIndex]) {
            const worksheet = this.worksheets.find(w => w.fileId === fileMapping.fileId);
            if (worksheet) {
              const columnIndex = worksheet.columns.indexOf(fileMapping.column);
              if (columnIndex >= 0 && columnIndex < fileData[rowIndex].data.length) {
                const cellValue = fileData[rowIndex].data[columnIndex];
                if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
                  combinedRowData[mappingIndex] = cellValue;
                }
              }
            }
          }
        });
      });
      
      // Only add row if it has some data
      if (combinedRowData.some(cell => cell !== '')) {
        combinedRowsData.push({
          data: combinedRowData,
          sourceFile: 'combined',
          sourceWorksheet: 'combined',
          originalRowIndex: rowIndex + 1
        });
      }
    }

    return combinedRowsData;
  }

  private detectDataType(values: string[]): string {
    if (values.length === 0) return 'unknown';
    
    const nonEmpty = values.filter(v => v.toString().trim() !== '');
    if (nonEmpty.length === 0) return 'empty';

    // Check if all values are numbers
    if (nonEmpty.every(v => !isNaN(Number(v)))) return 'number';
    
    // Check if all values are dates
    if (nonEmpty.every(v => !isNaN(Date.parse(v)))) return 'date';
    
    // Check if all values are boolean-like
    if (nonEmpty.every(v => ['true', 'false', 'yes', 'no', '1', '0'].includes(v.toLowerCase()))) return 'boolean';
    
    return 'text';
  }

  static generateCSVWithBOM(data: any[][]): Blob {
    // Convert data to CSV format
    const csvContent = data.map(row => 
      row.map(cell => {
        const cellValue = cell === null || cell === undefined ? '' : String(cell);
        // Escape quotes and wrap in quotes if contains comma, quote, or newline
        if (cellValue.includes(',') || cellValue.includes('"') || cellValue.includes('\n')) {
          return `"${cellValue.replace(/"/g, '""')}"`;
        }
        return cellValue;
      }).join(',')
    ).join('\n');

    // Add UTF-8 BOM for proper Unicode handling
    const BOM = '\uFEFF';
    const csvWithBOM = BOM + csvContent;

    return new Blob([csvWithBOM], { type: 'text/csv;charset=utf-8;' });
  }

  static downloadFile(blob: Blob, filename: string) {
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.URL.revokeObjectURL(url);
  }

  private combineDataByAlternativeKey(fileDataMap: Map<string, RowData[]>, keyColumnName: string): RowData[] {
    console.log('Using alternative key matching for column:', keyColumnName);
    
    // Create a map of key values to combined rows
    const keyRowMap = new Map<string, any[]>();

    // First pass: collect all unique key values from all files
    fileDataMap.forEach((fileRows, fileId) => {
      const worksheet = this.worksheets.find(w => w.fileId === fileId);
      if (!worksheet) return;

      const keyColumnIndex = worksheet.columns.indexOf(keyColumnName);
      if (keyColumnIndex < 0) {
        console.log(`Key column "${keyColumnName}" not found in worksheet for file ${fileId}`);
        return;
      }

      fileRows.forEach(row => {
        const keyValue = String(row.data[keyColumnIndex] || '').trim();
        if (keyValue && !keyRowMap.has(keyValue)) {
          keyRowMap.set(keyValue, new Array(this.mappings.length).fill(''));
        }
      });
    });

    console.log('Found unique key values:', keyRowMap.size);

    // Second pass: populate the combined rows
    fileDataMap.forEach((fileRows, fileId) => {
      const worksheet = this.worksheets.find(w => w.fileId === fileId);
      if (!worksheet) return;

      const keyColumnIndex = worksheet.columns.indexOf(keyColumnName);
      if (keyColumnIndex < 0) return;

      fileRows.forEach(row => {
        const keyValue = String(row.data[keyColumnIndex] || '').trim();
        if (!keyValue || !keyRowMap.has(keyValue)) return;

        const combinedRow = keyRowMap.get(keyValue)!;

        // Apply all column mappings for this file
        this.mappings.forEach((mapping, mappingIndex) => {
          const fileMapping = mapping.mappings.find(m => m.fileId === fileId);
          if (fileMapping) {
            const columnIndex = worksheet.columns.indexOf(fileMapping.column);
            if (columnIndex >= 0 && columnIndex < row.data.length) {
              const cellValue = row.data[columnIndex];
              if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
                combinedRow[mappingIndex] = cellValue;
              }
            }
          }
        });
      });
    });

    // Convert map to RowData array
    const combinedRowsData: RowData[] = [];
    keyRowMap.forEach((rowData, keyValue) => {
      if (rowData.some(cell => cell !== '')) {
        combinedRowsData.push({
          data: rowData,
          sourceFile: 'combined',
          sourceWorksheet: 'combined',
          originalRowIndex: combinedRowsData.length + 1
        });
      }
    });

    console.log('Alternative key matching produced rows:', combinedRowsData.length);
    return combinedRowsData;
  }

  private combineDataByWorksheetKeys(fileDataMap: Map<string, RowData[]>): RowData[] {
    console.log('Using worksheet-specific key columns');
    
    // Create a map of key values to combined rows
    const keyRowMap = new Map<string, any[]>();

    // First pass: collect all unique key values from all files using their specific key columns
    fileDataMap.forEach((fileRows, fileId) => {
      const worksheet = this.worksheets.find(w => w.fileId === fileId);
      if (!worksheet) return;

      // Use individual worksheet key column, fallback to global key column
      const keyColumnName = worksheet.keyColumn || this.keyColumn;
      if (!keyColumnName) {
        console.log(`No key column defined for worksheet in file ${fileId}`);
        return;
      }

      const keyColumnIndex = worksheet.columns.indexOf(keyColumnName);
      if (keyColumnIndex < 0) {
        console.log(`Key column "${keyColumnName}" not found in worksheet for file ${fileId}`);
        return;
      }

      console.log(`File ${fileId}: Using key column "${keyColumnName}" at index ${keyColumnIndex}`);

      fileRows.forEach(row => {
        const keyValue = String(row.data[keyColumnIndex] || '').trim();
        if (keyValue && !keyRowMap.has(keyValue)) {
          keyRowMap.set(keyValue, new Array(this.mappings.length).fill(''));
        }
      });
    });

    console.log('Found unique key values:', keyRowMap.size);

    // Second pass: populate the combined rows
    fileDataMap.forEach((fileRows, fileId) => {
      const worksheet = this.worksheets.find(w => w.fileId === fileId);
      if (!worksheet) return;

      const keyColumnName = worksheet.keyColumn || this.keyColumn;
      if (!keyColumnName) return;

      const keyColumnIndex = worksheet.columns.indexOf(keyColumnName);
      if (keyColumnIndex < 0) return;

      fileRows.forEach(row => {
        const keyValue = String(row.data[keyColumnIndex] || '').trim();
        if (!keyValue || !keyRowMap.has(keyValue)) return;

        const combinedRow = keyRowMap.get(keyValue)!;

        // Apply all column mappings for this file
        this.mappings.forEach((mapping, mappingIndex) => {
          const fileMapping = mapping.mappings.find(m => m.fileId === fileId);
          if (fileMapping) {
            const columnIndex = worksheet.columns.indexOf(fileMapping.column);
            if (columnIndex >= 0 && columnIndex < row.data.length) {
              const cellValue = row.data[columnIndex];
              if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
                combinedRow[mappingIndex] = cellValue;
              }
            }
          }
        });
      });
    });

    // Convert map to RowData array
    const combinedRowsData: RowData[] = [];
    keyRowMap.forEach((rowData, keyValue) => {
      if (rowData.some(cell => cell !== '')) {
        combinedRowsData.push({
          data: rowData,
          sourceFile: 'combined',
          sourceWorksheet: 'combined',
          originalRowIndex: combinedRowsData.length + 1
        });
      }
    });

    console.log('Worksheet key matching produced rows:', combinedRowsData.length);
    return combinedRowsData;
  }
}
