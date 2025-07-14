import React, { useState } from 'react';
import { FileSelection } from './excel-combiner/FileSelection';
import { WorksheetSelection } from './excel-combiner/WorksheetSelection';
import { ColumnPreview } from './excel-combiner/ColumnPreview';
import { ColumnMapping } from './excel-combiner/ColumnMapping';
import { Results } from './excel-combiner/Results';
import { Card } from './ui/card';
import { Badge } from './ui/badge';
import { ChevronRight, FileSpreadsheet } from 'lucide-react';
import { ExcelProcessor, ProcessingResults } from '../lib/excel-processor';
import { useToast } from '../hooks/use-toast';

export interface ExcelFile {
  id: string;
  name: string;
  file: File;
  worksheets: string[];
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

type Step = 'file-selection' | 'worksheet-selection' | 'column-preview' | 'column-mapping' | 'results';

export function ExcelCombiner() {
  const [currentStep, setCurrentStep] = useState<Step>('file-selection');
  const [selectedFiles, setSelectedFiles] = useState<ExcelFile[]>([]);
  const [selectedWorksheets, setSelectedWorksheets] = useState<WorksheetData[]>([]);
  const [keyColumn, setKeyColumn] = useState<string>('');
  const [columnMappings, setColumnMappings] = useState<ColumnMapping[]>([]);
  const [results, setResults] = useState<ProcessingResults | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const { toast } = useToast();

  const steps = [
    { id: 'file-selection', title: 'Select Files', icon: FileSpreadsheet },
    { id: 'worksheet-selection', title: 'Choose Worksheets', icon: FileSpreadsheet },
    { id: 'column-preview', title: 'Preview Columns', icon: FileSpreadsheet },
    { id: 'column-mapping', title: 'Map Columns', icon: FileSpreadsheet },
    { id: 'results', title: 'Results', icon: FileSpreadsheet },
  ];

  const currentStepIndex = steps.findIndex(step => step.id === currentStep);

  const handleNext = async () => {
    const nextIndex = currentStepIndex + 1;
    if (nextIndex < steps.length) {
      const nextStep = steps[nextIndex].id as Step;
      
      // If moving to results, process the data first
      if (nextStep === 'results') {
        await processExcelFiles();
      } else {
        setCurrentStep(nextStep);
      }
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
        description: "An error occurred while processing the Excel files. Please try again.",
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
    }
  };

  const handleStepClick = (stepId: string) => {
    const targetIndex = steps.findIndex(step => step.id === stepId);
    if (targetIndex <= currentStepIndex) {
      setCurrentStep(stepId as Step);
    }
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
              onFilesChange={setSelectedFiles}
              onNext={handleNext}
            />
          )}
          
          {currentStep === 'worksheet-selection' && (
            <WorksheetSelection
              selectedFiles={selectedFiles}
              selectedWorksheets={selectedWorksheets}
              onWorksheetsChange={setSelectedWorksheets}
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
              onWorksheetsChange={setSelectedWorksheets}
              onNext={handleNext}
              onBack={handleBack}
            />
          )}
          
          {currentStep === 'column-mapping' && (
            <ColumnMapping
              selectedFiles={selectedFiles}
              selectedWorksheets={selectedWorksheets}
              columnMappings={columnMappings}
              onMappingsChange={setColumnMappings}
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
              setCurrentStep('file-selection');
              setSelectedFiles([]);
              setSelectedWorksheets([]);
              setColumnMappings([]);
              setResults(null);
            }}
            worksheets={selectedWorksheets}
          />
          )}
        </Card>
      </div>
    </div>
  );
}