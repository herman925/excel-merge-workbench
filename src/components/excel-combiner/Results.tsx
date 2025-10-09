import React from 'react';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Badge } from '../ui/badge';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '../ui/table';
import { Download, ArrowLeft, CheckCircle, RotateCcw, FileText, AlertTriangle, TrendingUp, Eye } from 'lucide-react';
import { ProcessingResults, MergeLogEntry } from '../../lib/excel-processor';
import { ExcelProcessor } from '../../lib/excel-processor';
import { Accordion, AccordionContent, AccordionItem, AccordionTrigger } from '../ui/accordion';

interface ResultsProps {
  results: ProcessingResults | null;
  onBack: () => void;
  onStartOver: () => void;
  worksheets: WorksheetData[];
}

interface WorksheetData {
  fileId: string;
  worksheetName: string;
  headerRow: number;
  columns: string[];
}

export function Results({ results, onBack, onStartOver, worksheets }: ResultsProps) {
  if (!results) {
    return (
      <div className="p-8">
        <CardHeader className="text-center">
          <CardTitle className="text-2xl text-excel-primary">No Results Available</CardTitle>
          <CardDescription>Please process some files first.</CardDescription>
        </CardHeader>
      </div>
    );
  }
  const [logFilter, setLogFilter] = React.useState<'all' | 'conflicts'>('all');

  const handleDownload = () => {
    // Generate yyyymmdd format
    const today = new Date();
    const yyyymmdd = today.getFullYear().toString() + 
                    (today.getMonth() + 1).toString().padStart(2, '0') + 
                    today.getDate().toString().padStart(2, '0');
    
    // Determine most frequent worksheet name
    const worksheetNames = worksheets.map(w => w.worksheetName);
    const nameCounts = worksheetNames.reduce((acc, name) => {
      acc[name] = (acc[name] || 0) + 1;
      return acc;
    }, {} as Record<string, number>);
    
    // Find the name with highest frequency, or first one if tied
    let mostFrequentName = worksheetNames[0]; // fallback to first
    let maxCount = 0;
    
    for (const [name, count] of Object.entries(nameCounts)) {
      if (count > maxCount) {
        maxCount = count;
        mostFrequentName = name;
      }
    }
    
    const filename = `${yyyymmdd}_Combined_${mostFrequentName}.csv`;
    
    const blob = ExcelProcessor.generateCSVWithBOM(results.combinedData);
    ExcelProcessor.downloadFile(blob, filename);
  };

  return (
    <div className="p-8">
      <CardHeader className="text-center">
        <CardTitle className="text-2xl text-excel-primary">Processing Complete!</CardTitle>
        <CardDescription className="text-lg">
          Your Excel worksheets have been successfully combined
        </CardDescription>
      </CardHeader>

      <CardContent className="space-y-6">
        {/* Success Message */}
        <Card className="bg-gradient-to-r from-excel-accent-green/10 to-excel-accent-green/5 border-excel-accent-green/20">
          <CardContent className="p-6">
            <div className="flex items-center space-x-4">
              <div className="bg-excel-accent-green rounded-full p-3">
                <CheckCircle className="h-8 w-8 text-white" />
              </div>
              <div className="flex-1">
                <h2 className="text-2xl font-bold text-excel-accent-green mb-2">
                  Files Combined Successfully!
                </h2>
                <p className="text-lg text-muted-foreground">
                  {results.totalRowsProcessed} rows processed from {results.successfulFiles} files. Your data is ready for download.
                </p>
              </div>
            </div>
          </CardContent>
        </Card>

        {/* File Information */}
        <div className="grid md:grid-cols-2 gap-6">
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <FileText className="h-5 w-5 text-excel-primary" />
                <span>Output File</span>
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="flex justify-between">
                <span className="text-muted-foreground">Format:</span>
                <span className="font-medium">CSV (UTF-8 with BOM)</span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Columns:</span>
                <Badge variant="secondary">{results.combinedData[0]?.length || 0}</Badge>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Total Rows:</span>
                <Badge variant="secondary">{(results.combinedData.length - 1).toLocaleString()}</Badge>
              </div>
            </CardContent>
          </Card>

          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <TrendingUp className="h-5 w-5 text-excel-secondary" />
                <span>Processing Stats</span>
              </CardTitle>
            </CardHeader>
            <CardContent className="space-y-3">
              <div className="flex justify-between">
                <span className="text-muted-foreground">Files Processed:</span>
                <span className="font-medium">{results.successfulFiles}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Rows Processed:</span>
                <span className="font-medium">{results.totalRowsProcessed.toLocaleString()}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-muted-foreground">Duplicates Removed:</span>
                <Badge variant="secondary">{results.duplicatesRemoved}</Badge>
              </div>
            </CardContent>
          </Card>
        </div>

        {/* Data Preview */}
        {results.previewData.length > 0 && (
          <Card>
            <CardHeader>
              <CardTitle className="flex items-center space-x-2">
                <Eye className="h-5 w-5 text-excel-primary" />
                <span>Data Preview (First 10 Rows)</span>
              </CardTitle>
              <CardDescription>
                Preview of your combined data structure
              </CardDescription>
            </CardHeader>
            <CardContent>
              <div className="overflow-x-auto">
                <Table>
                  <TableHeader>
                    <TableRow>
                      {results.previewData[0]?.map((header, index) => (
                        <TableHead key={index} className="font-semibold text-excel-primary">
                          {header}
                        </TableHead>
                      ))}
                    </TableRow>
                  </TableHeader>
                  <TableBody>
                    {results.previewData.slice(1, 11).map((row, rowIndex) => (
                      <TableRow key={rowIndex}>
                        {row.map((cell, cellIndex) => (
                          <TableCell key={cellIndex} className="max-w-[200px] truncate">
                            {cell || '-'}
                          </TableCell>
                        ))}
                      </TableRow>
                    ))}
                  </TableBody>
                </Table>
              </div>
              {results.combinedData.length > 11 && (
                <p className="text-sm text-muted-foreground mt-3 text-center">
                  ... and {(results.combinedData.length - 11).toLocaleString()} more rows
                </p>
              )}
            </CardContent>
          </Card>
        )}

        {/* Download Button */}
        <Card className="bg-gradient-primary/10 border-excel-primary/20">
          <CardContent className="p-6 text-center">
            <h3 className="text-lg font-semibold mb-4">Ready to Download</h3>
            <Button
              onClick={handleDownload}
              size="lg"
              className="bg-gradient-primary hover:opacity-90 px-8 py-3"
            >
              <Download className="mr-2 h-5 w-5" />
              Download Combined CSV
            </Button>
            <p className="text-sm text-muted-foreground mt-3">
              UTF-8 CSV format with proper Unicode support
            </p>
          </CardContent>
        </Card>

        {/* Merge Log */}
        {results.mergeLog && results.mergeLog.length > 0 && (
          <Card className="border-purple-200 bg-purple-50 dark:bg-purple-900/20 dark:border-purple-800">
            <CardHeader>
              <CardTitle className="flex items-center space-x-2 text-purple-700 dark:text-purple-400">
                <FileText className="h-5 w-5" />
                <span>Column Merge Tracking</span>
              </CardTitle>
              <CardDescription className="text-purple-600 dark:text-purple-300">
                Detailed log showing which file's value was used for each cell
              </CardDescription>
            </CardHeader>
            <CardContent>
              <Accordion type="single" collapsible className="space-y-2">
                {results.mergeLog
                  .filter(entry => entry.conflictingValues.length > 0)
                  .slice(0, 20)
                  .map((entry, index) => (
                    <AccordionItem key={index} value={`item-${index}`} className="bg-white dark:bg-purple-900/30 rounded-lg border border-purple-200 dark:border-purple-700">
                      <AccordionTrigger className="px-4 py-3 hover:no-underline">
                        <div className="flex items-center justify-between w-full pr-4">
                          <span className="font-medium text-purple-800 dark:text-purple-200">
                            Row {entry.rowIndex} → {entry.outputColumn}
                          </span>
                          <Badge variant="destructive" className="ml-2">
                            {entry.conflictingValues.length} conflict{entry.conflictingValues.length > 1 ? 's' : ''}
                          </Badge>
                        </div>
                      </AccordionTrigger>
                      <AccordionContent className="px-4 pb-3 space-y-3">
                        <div className="p-3 bg-green-50 dark:bg-green-900/20 rounded border border-green-200 dark:border-green-700">
                          <div className="flex items-center space-x-2 mb-2">
                            <CheckCircle className="h-4 w-4 text-green-600" />
                            <span className="text-sm font-semibold text-green-700 dark:text-green-400">Value Used</span>
                          </div>
                          <div className="text-sm text-green-800 dark:text-green-300">
                            <span className="font-mono bg-white dark:bg-green-900/40 px-2 py-1 rounded">{entry.valueUsed || '(empty)'}</span>
                          </div>
                          <div className="text-xs text-green-600 dark:text-green-400 mt-2">
                            From: <span className="font-medium">{entry.sourceFile}</span> / {entry.sourceColumn}
                          </div>
                        </div>
                        
                        {entry.conflictingValues.map((conflict, cidx) => (
                          <div key={cidx} className="p-3 bg-red-50 dark:bg-red-900/20 rounded border border-red-200 dark:border-red-700">
                            <div className="flex items-center space-x-2 mb-2">
                              <AlertTriangle className="h-4 w-4 text-red-600" />
                              <span className="text-sm font-semibold text-red-700 dark:text-red-400">Conflicting Value (Ignored)</span>
                            </div>
                            <div className="text-sm text-red-800 dark:text-red-300">
                              <span className="font-mono bg-white dark:bg-red-900/40 px-2 py-1 rounded">{conflict.value || '(empty)'}</span>
                            </div>
                            <div className="text-xs text-red-600 dark:text-red-400 mt-2">
                              From: <span className="font-medium">{conflict.sourceFile}</span> / {conflict.sourceColumn}
                            </div>
                          </div>
                        ))}
                      </AccordionContent>
                    </AccordionItem>
                  ))}
              </Accordion>
              {results.mergeLog.filter(e => e.conflictingValues.length > 0).length > 20 && (
                <p className="text-sm text-purple-600 text-center mt-4">
                  ... and {results.mergeLog.filter(e => e.conflictingValues.length > 0).length - 20} more conflicts
                </p>
              )}
              {results.mergeLog.filter(e => e.conflictingValues.length > 0).length === 0 && (
                <div className="text-center py-4 text-purple-600 dark:text-purple-400">
                  <CheckCircle className="h-8 w-8 mx-auto mb-2" />
                  <p>No conflicts detected! All values merged cleanly.</p>
                </div>
              )}
            </CardContent>
          </Card>
        )}

        {/* Warnings/Issues */}
        {(results.unmappedColumns.length > 0 || results.duplicateRows.length > 0) && (
          <div className="grid md:grid-cols-2 gap-6">
            {/* Unmapped Columns Report */}
            {results.unmappedColumns.length > 0 && (
              <Card className="border-orange-200 bg-orange-50 dark:bg-orange-900/20 dark:border-orange-800">
                <CardHeader>
                  <CardTitle className="flex items-center space-x-2 text-orange-700 dark:text-orange-400">
                    <AlertTriangle className="h-5 w-5" />
                    <span>Unmapped Columns Report</span>
                  </CardTitle>
                  <CardDescription className="text-orange-600 dark:text-orange-300">
                    {results.unmappedColumns.length} columns were not mapped and excluded from output
                  </CardDescription>
                </CardHeader>
                <CardContent className="space-y-3">
                  {results.unmappedColumns.slice(0, 5).map((col, index) => (
                    <div key={index} className="p-3 bg-white dark:bg-orange-900/30 rounded-lg border border-orange-200 dark:border-orange-700">
                      <div className="flex items-center justify-between mb-2">
                        <span className="font-medium text-orange-800 dark:text-orange-200">{col.columnName}</span>
                        <Badge variant="outline" className="text-orange-700 border-orange-300">
                          {col.dataType}
                        </Badge>
                      </div>
                      <p className="text-sm text-orange-600 dark:text-orange-300">
                        From: {col.fileName} → {col.worksheetName}
                      </p>
                      {col.sampleValues.length > 0 && (
                        <div className="mt-2">
                          <p className="text-xs text-orange-500 dark:text-orange-400 mb-1">Sample values:</p>
                          <div className="flex flex-wrap gap-1">
                            {col.sampleValues.slice(0, 3).map((value, idx) => (
                              <Badge key={idx} variant="secondary" className="text-xs bg-orange-100 text-orange-700">
                                {value}
                              </Badge>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  ))}
                  {results.unmappedColumns.length > 5 && (
                    <p className="text-sm text-orange-600 text-center">
                      ... and {results.unmappedColumns.length - 5} more unmapped columns
                    </p>
                  )}
                </CardContent>
              </Card>
            )}

            {/* Duplicate Rows Report */}
            {results.duplicateRows.length > 0 && (
              <Card className="border-blue-200 bg-blue-50 dark:bg-blue-900/20 dark:border-blue-800">
                <CardHeader>
                  <CardTitle className="flex items-center space-x-2 text-blue-700 dark:text-blue-400">
                    <Eye className="h-5 w-5" />
                    <span>Duplicate Rows Report</span>
                  </CardTitle>
                  <CardDescription className="text-blue-600 dark:text-blue-300">
                    {results.duplicateRows.length} duplicate rows were detected and merged
                  </CardDescription>
                </CardHeader>
                <CardContent className="space-y-3">
                  {results.duplicateRows.slice(0, 5).map((dup, index) => (
                    <div key={index} className="p-3 bg-white dark:bg-blue-900/30 rounded-lg border border-blue-200 dark:border-blue-700">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-sm font-medium text-blue-800 dark:text-blue-200">Row {dup.rowIndex}</span>
                        <Badge variant="outline" className="text-blue-700 border-blue-300">
                          {dup.duplicateCount} duplicates
                        </Badge>
                      </div>
                      <div className="text-sm text-blue-600 dark:text-blue-300 mb-2">
                        <span className="font-medium">Data:</span> {dup.originalRow.slice(0, 3).join(' | ')}{dup.originalRow.length > 3 ? '...' : ''}
                      </div>
                      <div className="text-xs text-blue-500 dark:text-blue-400">
                        Found in: {dup.sourceFiles.join(', ')}
                      </div>
                    </div>
                  ))}
                  {results.duplicateRows.length > 5 && (
                    <p className="text-sm text-blue-600 text-center">
                      ... and {results.duplicateRows.length - 5} more duplicate groups
                    </p>
                  )}
                </CardContent>
              </Card>
            )}
          </div>
        )}

        {/* Navigation */}
        <div className="flex justify-between pt-4">
          <Button variant="outline" onClick={onBack}>
            <ArrowLeft className="mr-2 h-4 w-4" />
            Back to Mapping
          </Button>
          
          <Button 
            onClick={onStartOver}
            variant="outline"
            className="border-excel-primary text-excel-primary hover:bg-excel-primary hover:text-white"
          >
            <RotateCcw className="mr-2 h-4 w-4" />
            Start New Combination
          </Button>
        </div>
      </CardContent>
    </div>
  );
}