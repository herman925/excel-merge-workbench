import React, { useRef } from 'react';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Badge } from '../ui/badge';
import { Upload, X, FileSpreadsheet, AlertCircle } from 'lucide-react';
import { ExcelFile } from '../ExcelCombiner';

interface FileSelectionProps {
  selectedFiles: ExcelFile[];
  onFilesChange: (files: ExcelFile[]) => void;
  onNext: () => void;
}

export function FileSelection({ selectedFiles, onFilesChange, onNext }: FileSelectionProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(event.target.files || []);
    
    // Filter Excel files and limit to 5
    const excelFiles = files.filter(file => 
      file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    ).slice(0, 5 - selectedFiles.length);

    const newExcelFiles: ExcelFile[] = excelFiles.map((file, index) => ({
      id: `file-${selectedFiles.length + index + 1}`,
      name: file.name,
      file: file,
      worksheets: ['Sheet1', 'Sheet2', 'Data'] // Mock worksheets - in real app, these would be parsed
    }));

    onFilesChange([...selectedFiles, ...newExcelFiles]);
    
    // Clear input
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const removeFile = (fileId: string) => {
    onFilesChange(selectedFiles.filter(f => f.id !== fileId));
  };

  const canProceed = selectedFiles.length >= 2;

  return (
    <div className="p-8">
      <CardHeader className="text-center">
        <CardTitle className="text-2xl text-excel-primary">Select Excel Files</CardTitle>
        <CardDescription className="text-lg">
          Choose 2-5 Excel files to combine their worksheets
        </CardDescription>
      </CardHeader>

      <CardContent className="space-y-6">
        {/* File Upload Area */}
        <div className="border-2 border-dashed border-muted-foreground/25 rounded-lg p-8 text-center hover:border-excel-primary/50 transition-colors">
          <Upload className="mx-auto h-12 w-12 text-muted-foreground mb-4" />
          <h3 className="text-lg font-medium mb-2">Drop Excel files here</h3>
          <p className="text-muted-foreground mb-4">
            Or click to browse for .xlsx and .xls files
          </p>
          <Button 
            onClick={() => fileInputRef.current?.click()}
            className="bg-gradient-primary hover:opacity-90"
            disabled={selectedFiles.length >= 5}
          >
            <Upload className="mr-2 h-4 w-4" />
            Browse Files
          </Button>
          <input
            ref={fileInputRef}
            type="file"
            multiple
            accept=".xlsx,.xls"
            onChange={handleFileSelect}
            className="hidden"
          />
          <p className="text-sm text-muted-foreground mt-2">
            Maximum 5 files • Excel formats only
          </p>
        </div>

        {/* Selected Files */}
        {selectedFiles.length > 0 && (
          <div className="space-y-4">
            <h3 className="text-lg font-medium text-excel-primary">
              Selected Files ({selectedFiles.length}/5)
            </h3>
            <div className="grid gap-3">
              {selectedFiles.map((file) => (
                <Card key={file.id} className="p-4 bg-muted/30">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center space-x-3">
                      <FileSpreadsheet className="h-8 w-8 text-excel-secondary" />
                      <div>
                        <p className="font-medium">{file.name}</p>
                        <p className="text-sm text-muted-foreground">
                          {file.worksheets.length} worksheets found
                        </p>
                      </div>
                    </div>
                    <div className="flex items-center space-x-2">
                      <Badge variant="secondary">
                        {file.worksheets.length} sheets
                      </Badge>
                      <Button
                        variant="ghost"
                        size="sm"
                        onClick={() => removeFile(file.id)}
                        className="text-destructive hover:text-destructive"
                      >
                        <X className="h-4 w-4" />
                      </Button>
                    </div>
                  </div>
                </Card>
              ))}
            </div>
          </div>
        )}

        {/* Requirements */}
        <Card className="bg-muted/30 border-excel-accent-green/20">
          <CardContent className="p-4">
            <div className="flex items-start space-x-3">
              <AlertCircle className="h-5 w-5 text-excel-accent-green mt-0.5" />
              <div className="space-y-1">
                <p className="font-medium text-excel-accent-green">Requirements</p>
                <ul className="text-sm space-y-1 text-muted-foreground">
                  <li>• Minimum 2 files required to proceed</li>
                  <li>• Maximum 5 files can be selected</li>
                  <li>• Only .xlsx and .xls formats supported</li>
                  <li>• Files should contain similar data structures</li>
                </ul>
              </div>
            </div>
          </CardContent>
        </Card>

        {/* Navigation */}
        <div className="flex justify-end pt-4">
          <Button
            onClick={onNext}
            disabled={!canProceed}
            className="bg-gradient-primary hover:opacity-90 px-8"
          >
            Next: Select Worksheets
          </Button>
        </div>
      </CardContent>
    </div>
  );
}