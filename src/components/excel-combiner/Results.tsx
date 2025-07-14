import React from 'react';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Badge } from '../ui/badge';
import { 
  ArrowLeft, 
  Download, 
  CheckCircle, 
  RotateCcw, 
  FileText,
  BarChart3,
  Clock,
  Users
} from 'lucide-react';

interface ResultsProps {
  results: any;
  onBack: () => void;
  onStartOver: () => void;
}

export function Results({ results, onBack, onStartOver }: ResultsProps) {
  // Mock results data for demonstration
  const mockResults = {
    success: true,
    fileName: 'combined_worksheets.csv',
    totalRows: 12543,
    totalColumns: 8,
    filesProcessed: 3,
    processingTime: '2.3s',
    unmappedColumns: 2,
    duplicateRows: 45
  };

  const handleDownload = () => {
    // Mock download functionality
    const blob = new Blob(['Mock CSV content'], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = mockResults.fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="p-8">
      <CardHeader className="text-center">
        <div className="mx-auto w-16 h-16 bg-excel-accent-green/10 rounded-full flex items-center justify-center mb-4">
          <CheckCircle className="h-8 w-8 text-excel-accent-green" />
        </div>
        <CardTitle className="text-2xl text-excel-primary">Combination Complete!</CardTitle>
        <CardDescription className="text-lg">
          Your Excel worksheets have been successfully combined
        </CardDescription>
      </CardHeader>

      <CardContent className="space-y-6">
        {/* Success Summary */}
        <Card className="bg-excel-accent-green/5 border-excel-accent-green/20">
          <CardContent className="p-6">
            <div className="flex items-center space-x-3 mb-4">
              <FileText className="h-6 w-6 text-excel-accent-green" />
              <div>
                <h3 className="font-semibold text-excel-accent-green text-lg">
                  {mockResults.fileName}
                </h3>
                <p className="text-sm text-muted-foreground">
                  Ready for download • UTF-8 encoded
                </p>
              </div>
            </div>
            
            <div className="flex flex-wrap gap-2">
              <Badge variant="secondary" className="bg-excel-accent-green/10 text-excel-accent-green">
                {mockResults.totalRows.toLocaleString()} rows
              </Badge>
              <Badge variant="secondary" className="bg-excel-accent-green/10 text-excel-accent-green">
                {mockResults.totalColumns} columns
              </Badge>
              <Badge variant="secondary" className="bg-excel-accent-green/10 text-excel-accent-green">
                {mockResults.filesProcessed} files processed
              </Badge>
            </div>
          </CardContent>
        </Card>

        {/* Processing Statistics */}
        <div className="grid gap-4 md:grid-cols-2 lg:grid-cols-4">
          <Card>
            <CardContent className="p-4 text-center">
              <BarChart3 className="h-8 w-8 text-excel-primary mx-auto mb-2" />
              <div className="text-2xl font-bold text-excel-primary">
                {mockResults.totalRows.toLocaleString()}
              </div>
              <p className="text-sm text-muted-foreground">Total Rows</p>
            </CardContent>
          </Card>

          <Card>
            <CardContent className="p-4 text-center">
              <Users className="h-8 w-8 text-excel-secondary mx-auto mb-2" />
              <div className="text-2xl font-bold text-excel-secondary">
                {mockResults.filesProcessed}
              </div>
              <p className="text-sm text-muted-foreground">Files Combined</p>
            </CardContent>
          </Card>

          <Card>
            <CardContent className="p-4 text-center">
              <Clock className="h-8 w-8 text-excel-accent-green mx-auto mb-2" />
              <div className="text-2xl font-bold text-excel-accent-green">
                {mockResults.processingTime}
              </div>
              <p className="text-sm text-muted-foreground">Processing Time</p>
            </CardContent>
          </Card>

          <Card>
            <CardContent className="p-4 text-center">
              <FileText className="h-8 w-8 text-excel-accent-pink mx-auto mb-2" />
              <div className="text-2xl font-bold text-excel-accent-pink">
                {mockResults.totalColumns}
              </div>
              <p className="text-sm text-muted-foreground">Output Columns</p>
            </CardContent>
          </Card>
        </div>

        {/* Warnings/Issues */}
        {(mockResults.unmappedColumns > 0 || mockResults.duplicateRows > 0) && (
          <Card className="bg-amber-50 border-amber-200">
            <CardContent className="p-4">
              <h3 className="font-medium text-amber-800 mb-2">Processing Notes</h3>
              <ul className="text-sm text-amber-700 space-y-1">
                {mockResults.unmappedColumns > 0 && (
                  <li>• {mockResults.unmappedColumns} columns were not mapped and excluded from output</li>
                )}
                {mockResults.duplicateRows > 0 && (
                  <li>• {mockResults.duplicateRows} duplicate rows were detected and merged</li>
                )}
              </ul>
            </CardContent>
          </Card>
        )}

        {/* Download Section */}
        <Card className="bg-gradient-primary/5 border-excel-primary/20">
          <CardContent className="p-6 text-center">
            <h3 className="text-lg font-semibold text-excel-primary mb-4">
              Download Your Combined Data
            </h3>
            <Button 
              onClick={handleDownload}
              className="bg-gradient-primary hover:opacity-90 text-lg px-8 py-3"
              size="lg"
            >
              <Download className="mr-2 h-5 w-5" />
              Download CSV File
            </Button>
            <p className="text-sm text-muted-foreground mt-2">
              File will be saved in UTF-8 format for maximum compatibility
            </p>
          </CardContent>
        </Card>

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