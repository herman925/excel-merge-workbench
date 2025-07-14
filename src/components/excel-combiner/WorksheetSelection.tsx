import React from 'react';
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
  
  const handleWorksheetSelect = (fileId: string, worksheetName: string) => {
    const existing = selectedWorksheets.find(w => w.fileId === fileId);
    const file = selectedFiles.find(f => f.id === fileId);
    
    if (!file) return;
    
    const newWorksheet: WorksheetData = {
      fileId,
      worksheetName,
      headerRow: 1,
      columns: ['Column A', 'Column B', 'Column C', 'Column D'] // Mock columns
    };

    if (existing) {
      onWorksheetsChange(
        selectedWorksheets.map(w => w.fileId === fileId ? newWorksheet : w)
      );
    } else {
      onWorksheetsChange([...selectedWorksheets, newWorksheet]);
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