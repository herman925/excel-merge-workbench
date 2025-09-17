import React from 'react';
import * as XLSX from 'xlsx';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Badge } from '../ui/badge';
import { Input } from '../ui/input';
import { Label } from '../ui/label';
import { ArrowLeft, Eye, FileSpreadsheet, Settings, Loader2 } from 'lucide-react';
import { WorksheetData, ExcelFile } from '../ExcelCombiner';
import { readColumnHeaders } from '../../lib/excel-utils';

interface ColumnPreviewProps {
  selectedFiles: ExcelFile[];
  selectedWorksheets: WorksheetData[];
  onWorksheetsChange: (worksheets: WorksheetData[]) => void;
  onNext: () => void;
  onBack: () => void;
}

export function ColumnPreview({
  selectedFiles,
  selectedWorksheets,
  onWorksheetsChange,
  onNext,
  onBack
}: ColumnPreviewProps) {
  const [isUpdating, setIsUpdating] = React.useState<string>('');

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
          
          // Use the enhanced column header reading that handles merged cells
          const columns = readColumnHeaders(worksheet, headerRow);
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

  const updateHeaderRow = async (fileId: string, headerRow: number) => {
    const worksheet = selectedWorksheets.find(w => w.fileId === fileId);
    const file = selectedFiles.find(f => f.id === fileId);
    
    if (!worksheet || !file) return;
    
    setIsUpdating(fileId);
    
    try {
      const columns = await parseWorksheetColumns(file.file, worksheet.worksheetName, headerRow);
      
      onWorksheetsChange(
        selectedWorksheets.map(w => 
          w.fileId === fileId ? { ...w, headerRow, columns } : w
        )
      );
    } catch (error) {
      console.error('Failed to update header row:', error);
    } finally {
      setIsUpdating('');
    }
  };

  const getFileName = (fileId: string) => {
    const file = selectedFiles.find(f => f.id === fileId);
    return file ? file.name : fileId.replace('file-', 'File ');
  };

  const getTotalColumns = () => {
    return selectedWorksheets.reduce((sum, ws) => sum + ws.columns.length, 0);
  };

  const canProceed = selectedWorksheets.length > 0;

  return (
    <div className="p-8">
      <CardHeader className="text-center">
        <CardTitle className="text-2xl text-excel-primary">Preview Detected Columns</CardTitle>
        <CardDescription className="text-lg">
          Review and adjust header row settings for each worksheet before mapping
        </CardDescription>
      </CardHeader>

      <CardContent className="space-y-6">
        {/* Summary */}
        <Card className="bg-muted/30">
          <CardContent className="p-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-3">
                <Eye className="h-5 w-5 text-excel-primary" />
                <span className="font-medium">
                  {selectedWorksheets.length} worksheets • {getTotalColumns()} total columns detected
                </span>
              </div>
              <Badge variant="default">
                Ready for mapping
              </Badge>
            </div>
          </CardContent>
        </Card>

        {/* Worksheet Previews */}
        <div className="space-y-6">
          {selectedWorksheets.map((worksheet) => (
            <Card key={worksheet.fileId} className="border-excel-primary/20">
              <CardHeader className="pb-4">
                <div className="flex items-center justify-between">
                  <div className="flex items-center space-x-3">
                    <FileSpreadsheet className="h-6 w-6 text-excel-secondary" />
                    <div>
                      <CardTitle className="text-lg">{getFileName(worksheet.fileId)}</CardTitle>
                      <CardDescription>
                        Worksheet: <span className="font-medium">{worksheet.worksheetName}</span>
                      </CardDescription>
                    </div>
                  </div>
                  <Badge variant="secondary">
                    {worksheet.columns.length} columns
                  </Badge>
                </div>
              </CardHeader>

              <CardContent className="space-y-4">
                {/* Header Row Setting */}
                <div className="flex items-center space-x-4 p-4 bg-muted/30 rounded-lg">
                  <Settings className="h-5 w-5 text-excel-primary" />
                  <div className="flex-1">
                    <Label className="text-sm font-medium">Header Row</Label>
                    <p className="text-xs text-muted-foreground">
                      Which row contains the column headers?
                    </p>
                  </div>
                  <div className="flex items-center space-x-2">
                    <Label htmlFor={`header-${worksheet.fileId}`} className="text-sm">
                      Row:
                    </Label>
                    <div className="flex items-center space-x-2">
                      <Input
                        id={`header-${worksheet.fileId}`}
                        type="number"
                        min="1"
                        max="10"
                        value={worksheet.headerRow}
                        onChange={(e) => updateHeaderRow(worksheet.fileId, parseInt(e.target.value) || 1)}
                        className="w-20"
                        disabled={isUpdating === worksheet.fileId}
                      />
                      {isUpdating === worksheet.fileId && (
                        <Loader2 className="h-4 w-4 animate-spin text-excel-primary" />
                      )}
                    </div>
                  </div>
                </div>

                {/* Column Preview */}
                <div className="space-y-2">
                  <Label className="text-sm font-medium text-excel-primary">
                    Detected Columns (from row {worksheet.headerRow}):
                  </Label>
                  <div className="flex flex-wrap gap-2 p-4 bg-background rounded-lg border">
                    {worksheet.columns.length > 0 ? (
                      worksheet.columns.map((column, index) => (
                        <Badge key={index} variant="outline" className="text-xs">
                          {column || `Column ${index + 1}`}
                        </Badge>
                      ))
                    ) : (
                      <p className="text-sm text-muted-foreground italic">
                        No columns detected - please check header row setting
                      </p>
                    )}
                  </div>
                </div>

                {/* Column Statistics */}
                <div className="grid grid-cols-2 gap-4 text-sm">
                  <div>
                    <span className="text-muted-foreground">Total Columns:</span>
                    <span className="ml-2 font-medium">{worksheet.columns.length}</span>
                  </div>
                  <div>
                    <span className="text-muted-foreground">Empty Headers:</span>
                    <span className="ml-2 font-medium">
                      {worksheet.columns.filter(col => !col || col.trim() === '').length}
                    </span>
                  </div>
                </div>
              </CardContent>
            </Card>
          ))}
        </div>

        {/* Tips */}
        <Card className="bg-excel-accent-green/5 border-excel-accent-green/20">
          <CardContent className="p-4">
            <div className="flex items-start space-x-3">
              <Eye className="h-5 w-5 text-excel-accent-green mt-0.5" />
              <div className="space-y-1">
                <p className="font-medium text-excel-accent-green">Preview Tips</p>
                <ul className="text-sm space-y-1 text-muted-foreground">
                  <li>• Adjust header row if columns appear incorrect</li>
                  <li>• Empty or missing column names will be auto-generated</li>
                  <li>• All columns will be available for mapping in the next step</li>
                  <li>• You can change header row settings and see updated columns immediately</li>
                </ul>
              </div>
            </div>
          </CardContent>
        </Card>

        {/* Navigation */}
        <div className="flex justify-between pt-4">
          <Button variant="outline" onClick={onBack}>
            <ArrowLeft className="mr-2 h-4 w-4" />
            Back to Worksheets
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