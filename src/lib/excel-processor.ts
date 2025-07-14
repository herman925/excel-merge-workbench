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

  constructor(files: ExcelFile[], worksheets: WorksheetData[], mappings: ColumnMapping[]) {
    this.files = files;
    this.worksheets = worksheets;
    this.mappings = mappings;
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

    // Combine data from all files row by row (align by row index)
    const maxRowCount = Math.max(...Array.from(fileDataMap.values()).map(data => data.length));
    const combinedRowsData: RowData[] = [];
    
    for (let rowIndex = 0; rowIndex < maxRowCount; rowIndex++) {
      // For each row position, get data from all files that have this row
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
}
