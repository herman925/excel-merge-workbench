import React, { useRef } from 'react';
import * as XLSX from 'xlsx';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Badge } from '../ui/badge';
import { Upload, X, FileSpreadsheet, AlertCircle, Loader2, RotateCcw, RefreshCw } from 'lucide-react';
import { ExcelFile } from '../ExcelCombiner';

interface FileSelectionProps {
  selectedFiles: ExcelFile[];
  onFilesChange: (files: ExcelFile[]) => void;
  onNext: () => void;
}

export function FileSelection({ selectedFiles, onFilesChange, onNext }: FileSelectionProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [isLoading, setIsLoading] = React.useState(false);
  const [loadingFile, setLoadingFile] = React.useState<string>('');

  const parseExcelFile = async (file: File): Promise<string[]> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const workbook = XLSX.read(data, { type: 'array' });
          resolve(workbook.SheetNames);
        } catch (error) {
          console.error('Error parsing Excel file:', error);
          reject(error);
        }
      };
      
      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsArrayBuffer(file);
    });
  };

  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(event.target.files || []);
    
    // Filter Excel files and limit to 5
    const excelFiles = files.filter(file => 
      file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    ).slice(0, 5 - selectedFiles.length);

    if (excelFiles.length === 0) return;

    setIsLoading(true);
    
    // Parse each file to get real worksheet names
    const newExcelFiles: ExcelFile[] = [];
    
    for (let i = 0; i < excelFiles.length; i++) {
      const file = excelFiles[i];
      setLoadingFile(file.name);
      
      try {
        const worksheets = await parseExcelFile(file);
        // Filter out empty worksheet names
        const validWorksheets = worksheets.filter(ws => ws && ws.trim() !== '');
        
        newExcelFiles.push({
          id: `file-${selectedFiles.length + i + 1}`,
          name: file.name,
          file: file,
          worksheets: validWorksheets.length > 0 ? validWorksheets : ['Sheet1']
        });
      } catch (error) {
        console.error(`Failed to parse ${file.name}:`, error);
        // Fallback to default sheet names if parsing fails
        newExcelFiles.push({
          id: `file-${selectedFiles.length + i + 1}`,
          name: file.name,
          file: file,
          worksheets: ['Sheet1'] // Fallback
        });
      }
    }

    onFilesChange([...selectedFiles, ...newExcelFiles]);
    setIsLoading(false);
    setLoadingFile('');
    
    // Clear input
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const removeFile = (fileId: string) => {
    onFilesChange(selectedFiles.filter(f => f.id !== fileId));
  };

  const reloadFiles = async () => {
    if (selectedFiles.length === 0) return;
    
    setIsLoading(true);
    const reloadedFiles: ExcelFile[] = [];
    
    for (let i = 0; i < selectedFiles.length; i++) {
      const existingFile = selectedFiles[i];
      setLoadingFile(existingFile.name);
      
      try {
        const worksheets = await parseExcelFile(existingFile.file);
        const validWorksheets = worksheets.filter(ws => ws && ws.trim() !== '');
        
        reloadedFiles.push({
          ...existingFile,
          worksheets: validWorksheets.length > 0 ? validWorksheets : ['Sheet1']
        });
      } catch (error) {
        console.error(`Failed to reload ${existingFile.name}:`, error);
        // Keep existing data if reload fails
        reloadedFiles.push(existingFile);
      }
    }
    
    onFilesChange(reloadedFiles);
    setIsLoading(false);
    setLoadingFile('');
  };

  const reloadSingleFile = async (fileId: string) => {
    const fileToReload = selectedFiles.find(f => f.id === fileId);
    if (!fileToReload) return;

    setIsLoading(true);
    setLoadingFile(fileToReload.name);

    try {
      const worksheets = await parseExcelFile(fileToReload.file);
      const validWorksheets = worksheets.filter(ws => ws && ws.trim() !== '');
      
      const updatedFile = {
        ...fileToReload,
        worksheets: validWorksheets.length > 0 ? validWorksheets : ['Sheet1']
      };

      const updatedFiles = selectedFiles.map(f => 
        f.id === fileId ? updatedFile : f
      );
      
      onFilesChange(updatedFiles);
    } catch (error) {
      console.error(`Failed to reload ${fileToReload.name}:`, error);
    }

    setIsLoading(false);
    setLoadingFile('');
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
          
          {isLoading ? (
            <div className="space-y-3">
              <div className="flex items-center justify-center space-x-2">
                <Loader2 className="h-4 w-4 animate-spin" />
                <span className="text-sm">Processing {loadingFile}...</span>
              </div>
              <div className="w-64 mx-auto bg-muted rounded-full h-2">
                <div className="bg-excel-primary h-2 rounded-full animate-pulse w-1/2"></div>
              </div>
            </div>
          ) : (
            <Button 
              onClick={() => fileInputRef.current?.click()}
              className="bg-gradient-primary hover:opacity-90"
              disabled={selectedFiles.length >= 5}
            >
              <Upload className="mr-2 h-4 w-4" />
              Browse Files
            </Button>
          )}
          
          <input
            ref={fileInputRef}
            type="file"
            multiple
            accept=".xlsx,.xls"
            onChange={handleFileSelect}
            className="hidden"
            disabled={isLoading}
          />
          <p className="text-sm text-muted-foreground mt-2">
            Maximum 5 files • Excel formats only
          </p>
        </div>

        {/* Selected Files */}
         {selectedFiles.length > 0 && (
           <div className="space-y-4">
             <div className="flex items-center justify-between">
               <h3 className="text-lg font-medium text-excel-primary">
                 Selected Files ({selectedFiles.length}/5)
               </h3>
               <Button
                 variant="outline"
                 size="sm"
                 onClick={reloadFiles}
                 className="text-muted-foreground hover:text-foreground"
               >
                 <RotateCcw className="mr-2 h-4 w-4" />
                 Reload Files
               </Button>
             </div>
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
                         onClick={() => reloadSingleFile(file.id)}
                         className="text-muted-foreground hover:text-foreground"
                         disabled={isLoading}
                       >
                         <RefreshCw className="h-4 w-4" />
                       </Button>
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