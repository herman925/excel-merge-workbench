import React from 'react';
import * as XLSX from 'xlsx';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '../ui/select';
import { Input } from '../ui/input';
import { Label } from '../ui/label';
import { Badge } from '../ui/badge';
import { FileSpreadsheet, ArrowLeft, CheckCircle, Settings, Globe, Loader2 } from 'lucide-react';
import { ExcelFile, WorksheetData } from '../ExcelCombiner';

interface WorksheetSelectionProps {
  selectedFiles: ExcelFile[];
  selectedWorksheets: WorksheetData[];
  onWorksheetsChange: (worksheets: WorksheetData[]) => void;
  keyColumn?: string;
  onKeyColumnChange: (keyColumn: string) => void;
  onNext: () => void;
  onBack: () => void;
}

export function WorksheetSelection({
  selectedFiles,
  selectedWorksheets,
  onWorksheetsChange,
  keyColumn,
  onKeyColumnChange,
  onNext,
  onBack
}: WorksheetSelectionProps) {
  
  const [globalWorksheet, setGlobalWorksheet] = React.useState('');
  const [globalHeaderRow, setGlobalHeaderRow] = React.useState(1);
  const [isApplyingGlobal, setIsApplyingGlobal] = React.useState(false);
  
  const parseWorksheetColumns = async (file: File, worksheetName: string, headerRow: number = 1): Promise<string[]> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          const worksheet = workbook.Sheets[worksheetName];
          
          if (!worksheet) {
            resolve([]);
            return;
          }
          
          // Get the range of the worksheet
          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
          const columns: string[] = [];
          
          // Read the specified header row to get column names
          const headerRowIndex = headerRow - 1; // Convert to 0-based index
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: headerRowIndex, c: col });
            const cell = worksheet[cellAddress];
            if (cell && cell.v) {
              columns.push(String(cell.v));
            } else {
              columns.push(`Column ${String.fromCharCode(65 + col)}`);
            }
          }
          
          resolve(columns);
        } catch (error) {
          console.error('Error parsing worksheet columns:', error);
          resolve(['Column A', 'Column B', 'Column C']); // Fallback
        }
      };
      
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsArrayBuffer(file);
    });
  };

  const handleWorksheetSelect = async (fileId: string, worksheetName: string) => {
    const existing = selectedWorksheets.find(w => w.fileId === fileId);
    const file = selectedFiles.find(f => f.id === fileId);
    
    if (!file) return;
    
    // Handle "Not Selected" option
    if (worksheetName === 'NOT_SELECTED') {
      if (existing) {
        // Remove from selected worksheets
        onWorksheetsChange(selectedWorksheets.filter(w => w.fileId !== fileId));
      }
      return;
    }
    
    try {
      const columns = await parseWorksheetColumns(file.file, worksheetName, 1);
      
      const newWorksheet: WorksheetData = {
        fileId,
        worksheetName,
        headerRow: 1,
        columns: columns
      };

      if (existing) {
        onWorksheetsChange(
          selectedWorksheets.map(w => w.fileId === fileId ? newWorksheet : w)
        );
      } else {
        onWorksheetsChange([...selectedWorksheets, newWorksheet]);
      }
    } catch (error) {
      console.error('Failed to parse worksheet:', error);
      // Fallback with basic column names
      const newWorksheet: WorksheetData = {
        fileId,
        worksheetName,
        headerRow: 1,
        columns: ['Column A', 'Column B', 'Column C'] // Fallback
      };

      if (existing) {
        onWorksheetsChange(
          selectedWorksheets.map(w => w.fileId === fileId ? newWorksheet : w)
        );
      } else {
        onWorksheetsChange([...selectedWorksheets, newWorksheet]);
      }
    }
  };

  const updateHeaderRow = async (fileId: string, headerRow: number) => {
    const worksheet = selectedWorksheets.find(w => w.fileId === fileId);
    const file = selectedFiles.find(f => f.id === fileId);
    
    if (!worksheet || !file) return;
    
    try {
      const columns = await parseWorksheetColumns(file.file, worksheet.worksheetName, headerRow);
      
      onWorksheetsChange(
        selectedWorksheets.map(w => 
          w.fileId === fileId ? { ...w, headerRow, columns } : w
        )
      );
    } catch (error) {
      console.error('Failed to update header row:', error);
    }
  };

  const applyGlobalSettings = async () => {
    if (!globalWorksheet) return;
    
    setIsApplyingGlobal(true);
    
    const promises = selectedFiles.map(async (file) => {
      if (!file.worksheets.includes(globalWorksheet)) {
        // For files that don't have the global worksheet, set to "Not Selected"
        return null;
      }
      
      try {
        const columns = await parseWorksheetColumns(file.file, globalWorksheet, globalHeaderRow);
        
        return {
          fileId: file.id,
          worksheetName: globalWorksheet,
          headerRow: globalHeaderRow,
          columns: columns
        };
      } catch (error) {
        console.error(`Failed to parse ${file.name}:`, error);
        return {
          fileId: file.id,
          worksheetName: globalWorksheet,
          headerRow: globalHeaderRow,
          columns: ['Column A', 'Column B', 'Column C'] // Fallback
        };
      }
    });
    
    const results = await Promise.all(promises);
    const validResults = results.filter(Boolean) as WorksheetData[];
    
    // Remove selections for files that don't have the global worksheet
    const filesWithGlobalWorksheet = selectedFiles
      .filter(file => file.worksheets.includes(globalWorksheet))
      .map(file => file.id);
    
    const updatedWorksheets = selectedWorksheets.filter(w => 
      !filesWithGlobalWorksheet.includes(w.fileId)
    );
    
    // Add the new valid results
    validResults.forEach(newWorksheet => {
      const existingIndex = updatedWorksheets.findIndex(w => w.fileId === newWorksheet.fileId);
      if (existingIndex >= 0) {
        updatedWorksheets[existingIndex] = newWorksheet;
      } else {
        updatedWorksheets.push(newWorksheet);
      }
    });
    
    onWorksheetsChange(updatedWorksheets);
    setIsApplyingGlobal(false);
  };

  const getAvailableWorksheets = () => {
    const allWorksheets = selectedFiles.flatMap(file => file.worksheets);
    const worksheetCounts = allWorksheets.reduce((acc, ws) => {
      acc[ws] = (acc[ws] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
    
    // Return all unique worksheets with their counts
    return Object.entries(worksheetCounts)
      .map(([worksheet, count]) => ({ worksheet, count }));
  };

  const getWorksheetDisplayName = (worksheet: string, count: number) => {
    return `${worksheet} (${count}/${selectedFiles.length} files)`;
  };

  const getSelectedWorksheet = (fileId: string) => {
    return selectedWorksheets.find(w => w.fileId === fileId);
  };

  // Get common columns across all selected worksheets
  const getCommonColumns = () => {
    if (selectedWorksheets.length === 0) return [];
    
    const allColumns = selectedWorksheets.map(w => w.columns);
    if (allColumns.length === 0) return [];
    
    return allColumns[0].filter(column => 
      allColumns.every(worksheetColumns => worksheetColumns.includes(column))
    );
  };

  const canProceed = selectedWorksheets.length > 0;
  const availableGlobalWorksheets = getAvailableWorksheets();
  const commonColumns = getCommonColumns();

  return (
    <div className="p-8">
      <CardHeader className="text-center">
        <CardTitle className="text-2xl text-excel-primary">Select Worksheets</CardTitle>
        <CardDescription className="text-lg">
          Choose one worksheet from each Excel file to combine
        </CardDescription>
      </CardHeader>

      <CardContent className="space-y-6">
        {/* Global Controls */}
        {availableGlobalWorksheets.length > 0 && (
          <Card className="bg-gradient-secondary/10 border-excel-secondary/20">
            <CardContent className="p-6">
              <div className="flex items-center space-x-3 mb-4">
                <Globe className="h-5 w-5 text-excel-secondary" />
                <div>
                  <h3 className="font-semibold text-excel-secondary">Global Settings</h3>
                  <p className="text-sm text-muted-foreground">
                    Apply the same worksheet and header row to all files
                  </p>
                </div>
              </div>
              
              <div className="grid gap-4 md:grid-cols-3">
                <div className="space-y-2">
                  <Label className="text-sm font-medium">Worksheet Name</Label>
                  <Select value={globalWorksheet} onValueChange={setGlobalWorksheet}>
                    <SelectTrigger>
                      <SelectValue placeholder="Select worksheet..." />
                    </SelectTrigger>
                      <SelectContent className="bg-white z-50">
                        {availableGlobalWorksheets
                          .filter(({ worksheet }) => worksheet && typeof worksheet === 'string' && worksheet.trim().length > 0)
                          .map(({ worksheet, count }) => (
                            <SelectItem key={worksheet} value={worksheet}>
                              {getWorksheetDisplayName(worksheet, count)}
                            </SelectItem>
                          ))}
                      </SelectContent>
                  </Select>
                </div>
                
                <div className="space-y-2">
                  <Label className="text-sm font-medium">Header Row</Label>
                  <Input
                    type="number"
                    min="1"
                    max="10"
                    value={globalHeaderRow}
                    onChange={(e) => setGlobalHeaderRow(parseInt(e.target.value) || 1)}
                  />
                </div>
                
                <div className="flex items-end">
                  <Button
                    onClick={applyGlobalSettings}
                    disabled={!globalWorksheet || isApplyingGlobal}
                    className="w-full bg-excel-secondary hover:bg-excel-secondary/90"
                  >
                    {isApplyingGlobal ? (
                      <>
                        <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                        Applying...
                      </>
                    ) : (
                      'Apply to All'
                    )}
                  </Button>
                </div>
              </div>
              
              {globalWorksheet && (
                <p className="text-xs text-muted-foreground mt-3">
                  This will set "{globalWorksheet}" with header row {globalHeaderRow} for all files that contain this worksheet
                </p>
              )}
            </CardContent>
          </Card>
        )}

        {/* Key Column Selection */}
        {selectedWorksheets.length > 1 && commonColumns.length > 0 && (
          <Card className="bg-blue-50 border-blue-200 dark:bg-blue-900/20 dark:border-blue-800">
            <CardContent className="p-6">
              <div className="flex items-center space-x-3 mb-4">
                <Settings className="h-5 w-5 text-blue-600 dark:text-blue-400" />
                <div>
                  <h3 className="font-semibold text-blue-900 dark:text-blue-100">Key Column Matching</h3>
                  <p className="text-sm text-blue-700 dark:text-blue-300">
                    Select a common column to match rows across files (recommended for accurate data combination)
                  </p>
                </div>
              </div>
              
              <div className="grid gap-4 md:grid-cols-2">
                <div className="space-y-2">
                  <Label className="text-sm font-medium">Key Column</Label>
                  <Select value={keyColumn || '__none__'} onValueChange={(value) => onKeyColumnChange(value === '__none__' ? '' : value)}>
                    <SelectTrigger>
                      <SelectValue placeholder="Select key column..." />
                    </SelectTrigger>
                    <SelectContent className="bg-white z-50">
                      <SelectItem value="__none__">
                        <span className="text-muted-foreground italic">Use row position (not recommended)</span>
                      </SelectItem>
                      {commonColumns.map((column) => (
                        <SelectItem key={column} value={column}>
                          {column}
                        </SelectItem>
                      ))}
                    </SelectContent>
                  </Select>
                </div>
                
                <div className="flex items-end">
                  <div className="text-sm text-blue-700 dark:text-blue-300">
                    {keyColumn ? (
                      <div className="flex items-center space-x-2">
                        <CheckCircle className="h-4 w-4 text-green-600" />
                        <span>Rows will be matched by "{keyColumn}" values</span>
                      </div>
                    ) : (
                      <div className="text-orange-600 dark:text-orange-400">
                        Rows will be combined by position (may cause misalignment)
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </CardContent>
          </Card>
        )}

        {/* Individual File Selection */}
        <div className="space-y-4">
          <h3 className="text-lg font-medium text-excel-primary">Individual File Settings</h3>
          <div className="grid gap-6 md:grid-cols-2 lg:grid-cols-3">
          {selectedFiles.map((file) => {
            const selectedWorksheet = getSelectedWorksheet(file.id);
            const isSelected = !!selectedWorksheet;

            return (
              <Card key={file.id} className={`relative transition-all duration-200 ${
                isSelected ? 'border-excel-accent-green bg-excel-accent-green/5' : 'hover:shadow-md'
              }`}>
                {isSelected && (
                  <div className="absolute -top-2 -right-2 bg-excel-accent-green rounded-full p-1">
                    <CheckCircle className="h-4 w-4 text-white" />
                  </div>
                )}
                
                <CardHeader className="pb-3">
                  <div className="flex items-center space-x-3">
                    <FileSpreadsheet className="h-6 w-6 text-excel-secondary" />
                    <div className="min-w-0 flex-1">
                      <CardTitle className="text-base truncate">{file.name}</CardTitle>
                      <CardDescription className="text-sm">
                        {file.worksheets.length} worksheets available
                      </CardDescription>
                    </div>
                  </div>
                </CardHeader>

                <CardContent className="pt-0">
                  <div className="space-y-3">
                    <Select
                      value={selectedWorksheet?.worksheetName || 'NOT_SELECTED'}
                      onValueChange={(value) => handleWorksheetSelect(file.id, value)}
                    >
                      <SelectTrigger className="w-full">
                        <SelectValue placeholder="Select worksheet..." />
                      </SelectTrigger>
                      <SelectContent className="bg-white z-50">
                        <SelectItem value="NOT_SELECTED">
                          <span className="text-muted-foreground italic">Not Selected</span>
                        </SelectItem>
                        {file.worksheets
                          .filter(worksheet => worksheet && typeof worksheet === 'string' && worksheet.trim().length > 0)
                          .map((worksheet) => (
                            <SelectItem key={worksheet} value={worksheet}>
                              {worksheet}
                            </SelectItem>
                          ))}
                      </SelectContent>
                    </Select>

                    {selectedWorksheet && (
                      <div className="space-y-3">
                        <div className="flex items-center space-x-2 p-3 bg-muted/30 rounded-lg">
                          <Settings className="h-4 w-4 text-excel-primary" />
                          <Label className="text-sm flex-1">Header Row:</Label>
                          <Input
                            type="number"
                            min="1"
                            max="10"
                            value={selectedWorksheet.headerRow}
                            onChange={(e) => updateHeaderRow(file.id, parseInt(e.target.value) || 1)}
                            className="w-16 h-8"
                          />
                        </div>
                        
                        <div className="space-y-1">
                          <Badge variant="secondary" className="w-fit">
                            {selectedWorksheet.columns.length} columns detected
                          </Badge>
                          <p className="text-xs text-muted-foreground">
                            Click "Next" to preview all detected columns
                          </p>
                        </div>
                      </div>
                    )}
                  </div>
                </CardContent>
              </Card>
            );
          })}
          </div>
        </div>

        {/* Progress Summary */}
        {selectedWorksheets.length > 0 && (
          <Card className="bg-muted/30">
            <CardContent className="p-4">
              <div className="flex items-center justify-between">
                <div className="flex items-center space-x-3">
                  <CheckCircle className="h-5 w-5 text-excel-accent-green" />
                  <span className="font-medium">
                    {selectedWorksheets.length} of {selectedFiles.length} files selected
                  </span>
                </div>
                <Badge variant={canProceed ? "default" : "secondary"}>
                  {canProceed ? 'Ready to proceed' : 'At least one file must be selected'}
                </Badge>
              </div>
              
              {canProceed && (
                <div className="mt-3 text-sm text-muted-foreground">
                  Total columns to map: {selectedWorksheets.reduce((sum, w) => sum + w.columns.length, 0)}
                </div>
              )}
            </CardContent>
          </Card>
        )}

        {/* Navigation */}
        <div className="flex justify-between pt-4">
          <Button variant="outline" onClick={onBack}>
            <ArrowLeft className="mr-2 h-4 w-4" />
            Back to Files
          </Button>
          
          <Button
            onClick={onNext}
            disabled={!canProceed}
            className="bg-gradient-primary hover:opacity-90 px-8"
          >
            Next: Preview Columns
          </Button>
        </div>
      </CardContent>
    </div>
  );
}