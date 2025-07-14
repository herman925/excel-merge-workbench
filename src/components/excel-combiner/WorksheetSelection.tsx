import React from 'react';
import * as XLSX from 'xlsx';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '../ui/select';
import { Badge } from '../ui/badge';
import { FileSpreadsheet, ArrowLeft, CheckCircle } from 'lucide-react';
import { ExcelFile, WorksheetData } from '../ExcelCombiner';

interface WorksheetSelectionProps {
  selectedFiles: ExcelFile[];
  selectedWorksheets: WorksheetData[];
  onWorksheetsChange: (worksheets: WorksheetData[]) => void;
  onNext: () => void;
  onBack: () => void;
}

export function WorksheetSelection({
  selectedFiles,
  selectedWorksheets,
  onWorksheetsChange,
  onNext,
  onBack
}: WorksheetSelectionProps) {
  
  
  const parseWorksheetColumns = async (file: File, worksheetName: string): Promise<string[]> => {
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
          
          // Read the first row (header row) to get column names
          for (let col = range.s.c; col <= range.e.c; col++) {
            const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
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
    
    try {
      const columns = await parseWorksheetColumns(file.file, worksheetName);
      
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

  const getSelectedWorksheet = (fileId: string) => {
    return selectedWorksheets.find(w => w.fileId === fileId);
  };

  const canProceed = selectedWorksheets.length === selectedFiles.length;

  return (
    <div className="p-8">
      <CardHeader className="text-center">
        <CardTitle className="text-2xl text-excel-primary">Select Worksheets</CardTitle>
        <CardDescription className="text-lg">
          Choose one worksheet from each Excel file to combine
        </CardDescription>
      </CardHeader>

      <CardContent className="space-y-6">
        {/* File Selection Grid */}
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
                      value={selectedWorksheet?.worksheetName || ''}
                      onValueChange={(value) => handleWorksheetSelect(file.id, value)}
                    >
                      <SelectTrigger className="w-full">
                        <SelectValue placeholder="Select worksheet..." />
                      </SelectTrigger>
                      <SelectContent className="bg-white z-50">
                        {file.worksheets.map((worksheet) => (
                          <SelectItem key={worksheet} value={worksheet}>
                            {worksheet}
                          </SelectItem>
                        ))}
                      </SelectContent>
                    </Select>

                    {selectedWorksheet && (
                      <div className="space-y-2">
                        <Badge variant="secondary" className="w-fit">
                          {selectedWorksheet.columns.length} columns detected
                        </Badge>
                        <div className="text-xs text-muted-foreground">
                          Header row: {selectedWorksheet.headerRow}
                        </div>
                      </div>
                    )}
                  </div>
                </CardContent>
              </Card>
            );
          })}
        </div>

        {/* Progress Summary */}
        {selectedWorksheets.length > 0 && (
          <Card className="bg-muted/30">
            <CardContent className="p-4">
              <div className="flex items-center justify-between">
                <div className="flex items-center space-x-3">
                  <CheckCircle className="h-5 w-5 text-excel-accent-green" />
                  <span className="font-medium">
                    {selectedWorksheets.length} of {selectedFiles.length} worksheets selected
                  </span>
                </div>
                <Badge variant={canProceed ? "default" : "secondary"}>
                  {canProceed ? 'Ready to proceed' : 'Selection incomplete'}
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
            Next: Map Columns
          </Button>
        </div>
      </CardContent>
    </div>
  );
}