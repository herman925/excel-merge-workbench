import React, { useRef } from 'react';
import * as XLSX from 'xlsx';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Badge } from '../ui/badge';
import { Upload, X, FileSpreadsheet, AlertCircle, Loader2, RotateCcw, RefreshCw } from 'lucide-react';
import { ExcelFile } from '../ExcelCombiner';
import { useToast } from '../../hooks/use-toast';

interface FileSelectionProps {
  selectedFiles: ExcelFile[];
  onFilesChange: (files: ExcelFile[]) => void;
  onNext: () => void;
}

export function FileSelection({ selectedFiles, onFilesChange, onNext }: FileSelectionProps) {
  const fileInputRef = useRef<HTMLInputElement>(null);
  const replaceFileInputRef = useRef<HTMLInputElement>(null);
  const [isLoading, setIsLoading] = React.useState(false);
  const [loadingFile, setLoadingFile] = React.useState<string>('');
  const [fileIdToReplace, setFileIdToReplace] = React.useState<string | null>(null);
  const { toast } = useToast();

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

  const buildExcelFile = async (file: File, id: string): Promise<ExcelFile> => {
    try {
      const worksheets = await parseExcelFile(file);
      const validWorksheets = worksheets.filter(ws => ws && ws.trim() !== '');

      return {
        id,
        name: file.name,
        file,
        worksheets: validWorksheets.length > 0 ? validWorksheets : ['Sheet1'],
        readError: undefined,
      };
    } catch (error) {
      console.error(`Failed to parse ${file.name}:`, error);
      const msg = error instanceof Error ? error.message : String(error);
      const locked = /password|encrypt/i.test(msg);
      return {
        id,
        name: file.name,
        file,
        worksheets: ['Sheet1'],
        readError: locked
          ? 'This file is password-protected. Remove the password in Excel (File → Info → Protect Workbook → Encrypt), re-save, then re-select it here.'
          : 'This file could not be read. Please re-select it from disk.',
      };
    }
  };

  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(event.target.files || []);
    
    // Filter Excel files and limit to 10
    const excelFiles = files.filter(file =>
      file.name.endsWith('.xlsx') || file.name.endsWith('.xls')
    ).slice(0, 10 - selectedFiles.length);

    if (excelFiles.length === 0) return;

    setIsLoading(true);
    
    // Parse each file to get real worksheet names
    const newExcelFiles: ExcelFile[] = [];
    
    for (let i = 0; i < excelFiles.length; i++) {
      const file = excelFiles[i];
      setLoadingFile(file.name);

      newExcelFiles.push(await buildExcelFile(file, `file-${selectedFiles.length + i + 1}`));
    }

    onFilesChange([...selectedFiles, ...newExcelFiles]);
    setIsLoading(false);
    setLoadingFile('');
    
    // Clear input
    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const handleReplaceFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const replacementFile = event.target.files?.[0];

    if (!replacementFile || !fileIdToReplace) {
      return;
    }

    if (!(replacementFile.name.endsWith('.xlsx') || replacementFile.name.endsWith('.xls'))) {
      toast({
        title: 'Invalid file type',
        description: 'Please choose an .xlsx or .xls file.',
        variant: 'destructive'
      });
      return;
    }

    setIsLoading(true);
    setLoadingFile(replacementFile.name);

    const updatedFile = await buildExcelFile(replacementFile, fileIdToReplace);
    const updatedFiles = selectedFiles.map(file =>
      file.id === fileIdToReplace ? updatedFile : file
    );

    onFilesChange(updatedFiles);

    setIsLoading(false);
    setLoadingFile('');
    setFileIdToReplace(null);

    if (replaceFileInputRef.current) {
      replaceFileInputRef.current.value = '';
    }

    toast({
      title: 'File updated',
      description: `${updatedFile.name} was re-selected from disk and downstream worksheet data was refreshed automatically.`,
    });
  };

  const promptReplaceFile = (fileId: string) => {
    setFileIdToReplace(fileId);
    replaceFileInputRef.current?.click();
  };

  const removeFile = (fileId: string) => {
    console.log('Removing file with ID:', fileId);
    console.log('Current files:', selectedFiles.map(f => ({ id: f.id, name: f.name })));
    const filteredFiles = selectedFiles.filter(f => f.id !== fileId);
    console.log('Filtered files:', filteredFiles.map(f => ({ id: f.id, name: f.name })));
    onFilesChange(filteredFiles);
  };

  const reloadFiles = async () => {
    if (selectedFiles.length === 0) return;
    
    setIsLoading(true);
    const reloadedFiles: ExcelFile[] = [];
    const failedFiles: string[] = [];
    
    for (let i = 0; i < selectedFiles.length; i++) {
      const existingFile = selectedFiles[i];
      setLoadingFile(existingFile.name);
      
      try {
        const worksheets = await parseExcelFile(existingFile.file);
        const validWorksheets = worksheets.filter(ws => ws && ws.trim() !== '');
        
        reloadedFiles.push({
          ...existingFile,
          worksheets: validWorksheets.length > 0 ? validWorksheets : ['Sheet1'],
          readError: undefined,
        });
      } catch (error) {
        console.error(`Failed to reload ${existingFile.name}:`, error);
        failedFiles.push(existingFile.name);
        reloadedFiles.push({
          ...existingFile,
          readError: 'This file could not be re-read. Please use Re-select to load it again.',
        });
      }
    }
    
    onFilesChange(reloadedFiles);
    setIsLoading(false);
    setLoadingFile('');

    if (failedFiles.length > 0) {
      toast({
        title: 'Some files need re-selection',
        description: `${failedFiles.join(', ')} could not be re-read. If a workbook was edited outside the browser, use Re-select to load the updated file from disk.`,
        variant: 'destructive'
      });
    }
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
        worksheets: validWorksheets.length > 0 ? validWorksheets : ['Sheet1'],
        readError: undefined,
      };

      const updatedFiles = selectedFiles.map(f => 
        f.id === fileId ? updatedFile : f
      );
      
      onFilesChange(updatedFiles);
    } catch (error) {
      console.error(`Failed to reload ${fileToReload.name}:`, error);

      onFilesChange(selectedFiles.map(f => 
        f.id === fileId
          ? { ...f, readError: 'This file could not be re-read. Please use Re-select to load it again.' }
          : f
      ));

      toast({
        title: 'File needs re-selection',
        description: `${fileToReload.name} could not be re-read. If it was edited outside the browser, use Re-select to load the updated file from disk.`,
        variant: 'destructive'
      });
    }

    setIsLoading(false);
    setLoadingFile('');
  };

  const unreadableFiles = selectedFiles.filter(file => file.readError);
  const canProceed = selectedFiles.length >= 2 && unreadableFiles.length === 0;

  return (
    <div className="p-8">
      <CardHeader className="text-center">
        <CardTitle className="text-2xl text-excel-primary">Select Excel Files</CardTitle>
        <CardDescription className="text-lg">
          Choose 2-10 Excel files to combine their worksheets
        </CardDescription>
      </CardHeader>

      <CardContent className="space-y-6">
        {selectedFiles.length > 0 && (
          <Card className="border-excel-accent-green/30 bg-excel-accent-green/5">
            <CardContent className="p-4">
              <div className="flex items-start space-x-3">
                <FileSpreadsheet className="mt-0.5 h-5 w-5 text-excel-accent-green" />
                <div className="space-y-1">
                  <p className="font-medium text-excel-accent-green">Previously loaded files kept</p>
                  <p className="text-sm text-muted-foreground">
                    Your Excel files are still loaded. Worksheet selections, Apply to All, key columns, and mappings were reset, so you can safely start again from Step 1. If a workbook was edited outside the browser, use Re-select to reload that updated file from disk.
                  </p>
                </div>
              </div>
            </CardContent>
          </Card>
        )}

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
              disabled={selectedFiles.length >= 10}
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
            title="Select Excel files"
            onChange={handleFileSelect}
            className="hidden"
            disabled={isLoading}
          />
          <input
            ref={replaceFileInputRef}
            type="file"
            accept=".xlsx,.xls"
            title="Re-select an Excel file"
            onChange={handleReplaceFileSelect}
            className="hidden"
            disabled={isLoading}
          />
          <p className="text-sm text-muted-foreground mt-2">
            Maximum 10 files • Excel formats only
          </p>
        </div>

        {unreadableFiles.length > 0 && (
          <Card className="border-red-300 bg-red-50 dark:border-red-800 dark:bg-red-950/20">
            <CardContent className="p-4">
              <div className="flex items-start space-x-3">
                <AlertCircle className="mt-0.5 h-5 w-5 text-red-600 dark:text-red-400" />
                <div className="space-y-1">
                  <p className="font-medium text-red-700 dark:text-red-400">Re-select unreadable files before continuing</p>
                  <p className="text-sm text-red-600 dark:text-red-300">
                    {unreadableFiles.map(file => file.name).join(', ')} cannot be used until they are re-selected from disk.
                  </p>
                </div>
              </div>
            </CardContent>
          </Card>
        )}

        {/* Selected Files */}
         {selectedFiles.length > 0 && (
           <div className="space-y-4">
             <div className="flex items-center justify-between">
               <h3 className="text-lg font-medium text-excel-primary">
                 Selected Files ({selectedFiles.length}/10)
               </h3>
               <Button
                 variant="outline"
                 size="sm"
                 onClick={reloadFiles}
                 className="text-muted-foreground hover:text-foreground"
               >
                 <RotateCcw className="mr-2 h-4 w-4" />
                 Re-parse Files
               </Button>
             </div>
            <div className="grid gap-3">
              {selectedFiles.map((file) => (
                <Card key={file.id} className={`p-4 ${file.readError ? 'border-red-300 bg-red-50 dark:border-red-800 dark:bg-red-950/20' : 'bg-muted/30'}`}>
                  <div className="flex items-center justify-between">
                    <div className="flex items-center space-x-3">
                      <FileSpreadsheet className="h-8 w-8 text-excel-secondary" />
                      <div>
                        <p className="font-medium">{file.name}</p>
                        <p className="text-sm text-muted-foreground">
                          {file.readError || `${file.worksheets.length} worksheets found`}
                        </p>
                      </div>
                    </div>
                     <div className="flex items-center space-x-2">
                       <Badge variant={file.readError ? 'destructive' : 'secondary'}>
                         {file.readError ? 'Needs re-select' : `${file.worksheets.length} sheets`}
                       </Badge>
                       <Button
                         variant="outline"
                         size="sm"
                         onClick={() => promptReplaceFile(file.id)}
                         disabled={isLoading}
                       >
                         <Upload className="mr-2 h-4 w-4" />
                         Re-select
                       </Button>
                       <Button
                         variant="ghost"
                         size="sm"
                         onClick={() => reloadSingleFile(file.id)}
                         className="text-muted-foreground hover:text-foreground"
                         disabled={isLoading}
                         title="Re-parse the currently loaded file copy"
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
                  <li>• Maximum 10 files can be selected</li>
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
            title={unreadableFiles.length > 0 ? 'Re-select unreadable files before continuing' : undefined}
          >
            Next: Select Worksheets
          </Button>
        </div>
      </CardContent>
    </div>
  );
}