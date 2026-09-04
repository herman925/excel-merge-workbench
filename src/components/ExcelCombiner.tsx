import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { FileSelection } from './excel-combiner/FileSelection';
import { WorksheetSelection } from './excel-combiner/WorksheetSelection';
import { ColumnPreview } from './excel-combiner/ColumnPreview';
import { ColumnMapping } from './excel-combiner/ColumnMapping';
import { RearrangeColumns } from './excel-combiner/RearrangeColumns';
import { Results } from './excel-combiner/Results';
import { ConfigManager, ImportResult, PendingItem } from './excel-combiner/ConfigManager';
import { Card } from './ui/card';
import { Badge } from './ui/badge';
import { ChevronRight, FileSpreadsheet, AlertTriangle } from 'lucide-react';
import { ExcelProcessor, ProcessingResults } from '../lib/excel-processor';
import { useToast } from '../hooks/use-toast';

export interface ExcelFile {
  id: string;
  name: string;
  file: File;
  worksheets: string[];
  readError?: string;
}

export interface WorksheetData {
  fileId: string;
  worksheetName: string;
  headerRow: number;
  columns: string[];
  keyColumn?: string;
}

export interface ColumnMapping {
  outputColumn: string;
  mappings: { fileId: string; column: string }[];
}

type Step = 'file-selection' | 'worksheet-selection' | 'column-preview' | 'column-mapping' | 'rearrange-columns' | 'results';

export function ExcelCombiner() {
  const [currentStep, setCurrentStep] = useState<Step>('file-selection');
  const [selectedFiles, setSelectedFiles] = useState<ExcelFile[]>([]);
  const [selectedWorksheets, setSelectedWorksheets] = useState<WorksheetData[]>([]);
  const [keyColumn, setKeyColumn] = useState<string>('');
  const [columnMappings, setColumnMappings] = useState<ColumnMapping[]>([]);
  const [allowIncompleteMappings, setAllowIncompleteMappings] = useState(false);
  const [allowDoubleMapping, setAllowDoubleMapping] = useState(false);
  const [results, setResults] = useState<ProcessingResults | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [pending, setPending] = useState<PendingItem[]>([]);
  const { toast } = useToast();

  const steps = [
    { id: 'file-selection', title: 'Select Files', icon: FileSpreadsheet },
    { id: 'worksheet-selection', title: 'Choose Worksheets', icon: FileSpreadsheet },
    { id: 'column-preview', title: 'Preview Columns', icon: FileSpreadsheet },
    { id: 'column-mapping', title: 'Map Columns', icon: FileSpreadsheet },
    { id: 'rearrange-columns', title: 'Rearrange Columns', icon: FileSpreadsheet },
    { id: 'results', title: 'Results', icon: FileSpreadsheet },
  ];

  const currentStepIndex = steps.findIndex(step => step.id === currentStep);

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

          const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
          const columns: string[] = [];
          const headerRowIndex = headerRow - 1;

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
          reject(error);
        }
      };

      reader.onerror = () => reject(new Error('Failed to read file'));
      reader.readAsArrayBuffer(file);
    });
  };

  const handleFilesChange = async (files: ExcelFile[]) => {
    const validFileIds = new Set(files.filter(file => !file.readError).map(file => file.id));
    const previousFilesById = new Map(selectedFiles.map(file => [file.id, file]));
    const changedFileIds = new Set(
      files
        .filter((file) => previousFilesById.get(file.id)?.file !== file.file)
        .map((file) => file.id)
    );

    setSelectedFiles(files);
    setResults(null);
    setPending((prev) => prev.filter((p) => p.step !== 'file-selection'));

    const retainedWorksheets = selectedWorksheets.filter((worksheet) => validFileIds.has(worksheet.fileId));
    const reparsedWorksheets = await Promise.all(
      retainedWorksheets.map(async (worksheet) => {
        if (!changedFileIds.has(worksheet.fileId)) {
          return worksheet;
        }

        const file = files.find((candidate) => candidate.id === worksheet.fileId);
        if (!file || file.readError || !file.worksheets.includes(worksheet.worksheetName)) {
          return null;
        }

        try {
          const columns = await parseWorksheetColumns(file.file, worksheet.worksheetName, worksheet.headerRow);
          return {
            ...worksheet,
            columns,
            keyColumn: worksheet.keyColumn && columns.includes(worksheet.keyColumn)
              ? worksheet.keyColumn
              : undefined,
          };
        } catch {
          return null;
        }
      })
    );

    const nextWorksheets = reparsedWorksheets.filter((worksheet): worksheet is WorksheetData => !!worksheet);
    setSelectedWorksheets(nextWorksheets);

    if (nextWorksheets.length === 0) {
      setKeyColumn('');
    } else if (keyColumn) {
      const keyStillExists = nextWorksheets.some((worksheet) => worksheet.keyColumn === keyColumn);
      if (!keyStillExists) {
        setKeyColumn('');
      }
    }

    const worksheetMap = new Map(nextWorksheets.map((worksheet) => [worksheet.fileId, worksheet]));
    setColumnMappings(prev => prev.map(mapping => ({
      ...mapping,
      mappings: mapping.mappings.filter(fileMapping => {
        const worksheet = worksheetMap.get(fileMapping.fileId);
        return !!worksheet && worksheet.columns.includes(fileMapping.column);
      })
    })));

    if (files.length === 0) {
      setAllowIncompleteMappings(false);
    }
  };

  const handleFileReadError = (fileId: string, message: string) => {
    setSelectedFiles((prev) => prev.map((file) =>
      file.id === fileId ? { ...file, readError: message } : file
    ));
    setSelectedWorksheets((prev) => prev.filter((worksheet) => worksheet.fileId !== fileId));
    setColumnMappings((prev) => prev.map((mapping) => ({
      ...mapping,
      mappings: mapping.mappings.filter((fileMapping) => fileMapping.fileId !== fileId),
    })));
    setResults(null);
  };

  const normalizeColumnMappings = (mappings: ColumnMapping[], shouldAllowDoubleMapping: boolean) => {
    if (shouldAllowDoubleMapping) {
      return mappings;
    }

    const usedMappings = new Set<string>();

    return mappings.map((mapping) => ({
      ...mapping,
      mappings: mapping.mappings.filter((fileMapping) => {
        const key = `${fileMapping.fileId}::${fileMapping.column}`;

        if (!fileMapping.column?.trim()) {
          return false;
        }

        if (usedMappings.has(key)) {
          return false;
        }

        usedMappings.add(key);
        return true;
      }),
    }));
  };

  const handleMappingsChange = (mappings: ColumnMapping[]) => {
    setColumnMappings(normalizeColumnMappings(mappings, allowDoubleMapping));
    setResults(null);
    setPending((prev) => prev.filter((p) => p.step !== 'column-mapping'));
  };

  const handleAllowDoubleMappingChange = (value: boolean) => {
    setAllowDoubleMapping(value);
    setColumnMappings((prev) => normalizeColumnMappings(prev, value));
    setResults(null);
  };

  const handleNext = async () => {
    const unreadableFiles = selectedFiles.filter((file) => file.readError);
    if (unreadableFiles.length > 0) {
      setCurrentStep('file-selection');
      toast({
        title: 'Re-select unreadable files',
        description: `${unreadableFiles.map((file) => file.name).join(', ')} must be re-selected before continuing.`,
        variant: 'destructive'
      });
      window.scrollTo({ top: 0, behavior: 'smooth' });
      return;
    }

    const nextIndex = currentStepIndex + 1;
    if (nextIndex < steps.length) {
      const nextStep = steps[nextIndex].id as Step;
      
      // If moving to results, process the data first
      if (nextStep === 'results') {
        await processExcelFiles();
      } else {
        setCurrentStep(nextStep);
      }
      
      // Scroll to top of page
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  };

  const processExcelFiles = async () => {
    if (selectedFiles.length === 0 || selectedWorksheets.length === 0 || columnMappings.length === 0) {
      toast({
        title: "Missing Data",
        description: "Please ensure all files, worksheets, and column mappings are configured.",
        variant: "destructive"
      });
      return;
    }

    setIsProcessing(true);
    
    try {
      const processor = new ExcelProcessor(selectedFiles, selectedWorksheets, columnMappings, keyColumn);
      const processingResults = await processor.processFiles();
      
      setResults(processingResults);
      setCurrentStep('results');
      
      toast({
        title: "Processing Complete",
        description: `Successfully combined ${processingResults.totalRowsProcessed} rows from ${processingResults.successfulFiles} files.`,
      });
    } catch (error) {
      console.error('Error processing Excel files:', error);
      toast({
        title: "Processing Failed",
        description: error instanceof Error ? error.message : "An error occurred while processing the Excel files. Please try again.",
        variant: "destructive"
      });
    } finally {
      setIsProcessing(false);
    }
  };

  const handleBack = () => {
    const prevIndex = currentStepIndex - 1;
    if (prevIndex >= 0) {
      setCurrentStep(steps[prevIndex].id as Step);
      // Scroll to top of page
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  };

  const handleWorksheetsChange = (worksheets: WorksheetData[]) => {
    setSelectedWorksheets(worksheets);
    setPending((prev) => prev.filter((p) => p.step !== 'worksheet-selection'));
  };

  const handleStepClick = (stepId: string) => {
    const targetIndex = steps.findIndex(step => step.id === stepId);
    if (targetIndex <= currentStepIndex) {
      setCurrentStep(stepId as Step);
    }
  };

  const handleApplyImport = (result: ImportResult) => {
    setSelectedFiles(result.files);
    setSelectedWorksheets(result.worksheets);
    setColumnMappings(result.columnMappings);
    setKeyColumn(result.keyColumn);
    setAllowIncompleteMappings(
      result.allowIncompleteMappings ||
      result.pending.some(p => p.step === 'column-mapping')
    );
    setAllowDoubleMapping(result.allowDoubleMapping);
    setPending(result.pending);
    setResults(null);

    // Advance to the furthest step the import could populate, or back to step 1 if files are missing.
    if (result.files.length === 0) {
      setCurrentStep('file-selection');
    } else if (result.worksheets.length === 0) {
      setCurrentStep('worksheet-selection');
    } else {
      setCurrentStep('column-mapping');
    }
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-background to-muted/30">
      <div className="container mx-auto px-4 py-8">
        {/* Header */}
        <div className="text-center mb-8">
          <h1 className="text-4xl font-bold bg-gradient-hero bg-clip-text text-transparent mb-4">
            Excel Worksheet Combiner
          </h1>
          <p className="text-lg text-muted-foreground max-w-2xl mx-auto">
            Combine worksheets from multiple Excel files into a single CSV with intelligent column mapping
          </p>
          <div className="flex justify-center mt-4">
            <ConfigManager
              files={selectedFiles}
              worksheets={selectedWorksheets}
              columnMappings={columnMappings}
              keyColumn={keyColumn}
              allowIncompleteMappings={allowIncompleteMappings}
              allowDoubleMapping={allowDoubleMapping}
              onApplyImport={handleApplyImport}
            />
          </div>
          {pending.length > 0 && (
            <div className="mt-3 flex flex-wrap justify-center gap-2">
              {pending.map((item, i) => (
                <Badge key={i} variant="outline" className="border-orange-300 text-orange-700 dark:text-orange-300">
                  <AlertTriangle className="h-3 w-3 mr-1" /> Pending: {item.message}
                </Badge>
              ))}
            </div>
          )}
        </div>

        {/* Progress Steps */}
        <Card className="mb-8 p-6 shadow-card">
          <div className="flex items-center justify-between">
            {steps.map((step, index) => {
              const isActive = step.id === currentStep;
              const isCompleted = index < currentStepIndex;
              const isAccessible = index <= currentStepIndex;

              return (
                <div key={step.id} className="flex items-center">
                  <div 
                    className={`flex flex-col items-center cursor-pointer transition-all duration-200 ${
                      isAccessible ? 'hover:scale-105' : 'cursor-not-allowed opacity-50'
                    }`}
                    onClick={() => isAccessible && handleStepClick(step.id)}
                  >
                    <div className={`
                      w-12 h-12 rounded-full flex items-center justify-center mb-2 transition-all duration-200
                      ${isActive ? 'bg-gradient-primary text-white shadow-primary' : 
                        isCompleted ? 'bg-excel-accent-green text-white' :
                        'bg-muted text-muted-foreground'}
                    `}>
                      <step.icon size={20} />
                    </div>
                    <span className={`text-sm font-medium ${
                      isActive ? 'text-excel-primary' : 
                      isCompleted ? 'text-excel-accent-green' :
                      'text-muted-foreground'
                    }`}>
                      {step.title}
                    </span>
                    {isActive && (
                      <Badge variant="secondary" className="mt-1 text-xs">
                        Current
                      </Badge>
                    )}
                  </div>
                  {index < steps.length - 1 && (
                    <ChevronRight className={`mx-4 ${
                      index < currentStepIndex ? 'text-excel-accent-green' : 'text-muted-foreground'
                    }`} size={20} />
                  )}
                </div>
              );
            })}
          </div>
        </Card>

        {/* Step Content */}
        <Card className="shadow-card">
          {currentStep === 'file-selection' && (
            <FileSelection 
              selectedFiles={selectedFiles}
              onFilesChange={handleFilesChange}
              onNext={handleNext}
            />
          )}
          
          {currentStep === 'worksheet-selection' && (
            <WorksheetSelection
              selectedFiles={selectedFiles}
              selectedWorksheets={selectedWorksheets}
              onWorksheetsChange={handleWorksheetsChange}
              onFileReadError={handleFileReadError}
              keyColumn={keyColumn}
              onKeyColumnChange={setKeyColumn}
              onNext={handleNext}
              onBack={handleBack}
            />
          )}
          
          
          {currentStep === 'column-preview' && (
            <ColumnPreview
              selectedFiles={selectedFiles}
              selectedWorksheets={selectedWorksheets}
              onWorksheetsChange={handleWorksheetsChange}
              onFileReadError={handleFileReadError}
              onNext={handleNext}
              onBack={handleBack}
            />
          )}
          
          {currentStep === 'column-mapping' && (
            <ColumnMapping
              selectedFiles={selectedFiles}
              selectedWorksheets={selectedWorksheets}
              columnMappings={columnMappings}
              onMappingsChange={handleMappingsChange}
              allowIncompleteMappings={allowIncompleteMappings}
              onAllowIncompleteMappingsChange={setAllowIncompleteMappings}
              allowDoubleMapping={allowDoubleMapping}
              onAllowDoubleMappingChange={handleAllowDoubleMappingChange}
              onNext={handleNext}
              onBack={handleBack}
              isProcessing={isProcessing}
            />
          )}

          {currentStep === 'rearrange-columns' && (
            <RearrangeColumns
              columnMappings={columnMappings}
              selectedWorksheets={selectedWorksheets}
              onMappingsChange={handleMappingsChange}
              onNext={handleNext}
              onBack={handleBack}
              isProcessing={isProcessing}
            />
          )}
          
          {currentStep === 'results' && (
          <Results 
            results={results} 
            onBack={handleBack} 
            onStartOver={() => {
              setSelectedWorksheets([]);
              setKeyColumn('');
              setColumnMappings([]);
              setAllowIncompleteMappings(false);
              setAllowDoubleMapping(false);
              setCurrentStep('file-selection');
              setResults(null);
              setPending([]);
              window.scrollTo({ top: 0, behavior: 'smooth' });
            }}
            worksheets={selectedWorksheets}
          />
          )}
        </Card>
      </div>
    </div>
  );
}