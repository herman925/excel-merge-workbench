import React, { useState } from 'react';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Input } from '../ui/input';
import { Label } from '../ui/label';
import { Badge } from '../ui/badge';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '../ui/select';
import { ArrowLeft, Wand2, Plus, X, Link, AlertTriangle, CheckCircle, XCircle } from 'lucide-react';
import { WorksheetData, ColumnMapping as ColumnMappingType, ExcelFile } from '../ExcelCombiner';

interface ColumnMappingProps {
  selectedFiles: ExcelFile[];
  selectedWorksheets: WorksheetData[];
  columnMappings: ColumnMappingType[];
  onMappingsChange: (mappings: ColumnMappingType[]) => void;
  onNext: () => void;
  onBack: () => void;
}

export function ColumnMapping({
  selectedFiles,
  selectedWorksheets,
  columnMappings,
  onMappingsChange,
  onNext,
  onBack
}: ColumnMappingProps) {
  const [newColumnName, setNewColumnName] = useState('');

  const autoMap = () => {
    // Simple auto-mapping logic based on column name similarity
    const allColumns = selectedWorksheets.flatMap(ws => 
      ws.columns.map(col => ({ fileId: ws.fileId, column: col }))
    );
    
    const uniqueColumnNames = Array.from(new Set(allColumns.map(c => c.column.toLowerCase())));
    
    const newMappings: ColumnMappingType[] = uniqueColumnNames.map(colName => ({
      outputColumn: colName.charAt(0).toUpperCase() + colName.slice(1),
      mappings: allColumns.filter(c => c.column.toLowerCase() === colName)
    }));

    onMappingsChange(newMappings);
  };

  const addNewMapping = () => {
    if (!newColumnName.trim()) return;
    
    const newMapping: ColumnMappingType = {
      outputColumn: newColumnName.trim(),
      mappings: []
    };

    onMappingsChange([...columnMappings, newMapping]);
    setNewColumnName('');
  };

  const removeMapping = (index: number) => {
    onMappingsChange(columnMappings.filter((_, i) => i !== index));
  };

  const updateMappingColumn = (mappingIndex: number, fileId: string, column: string) => {
    const updatedMappings = columnMappings.map((mapping, index) => {
      if (index === mappingIndex) {
        const existingMapping = mapping.mappings.find(m => m.fileId === fileId);
        
        if (existingMapping) {
          return {
            ...mapping,
            mappings: mapping.mappings.map(m => 
              m.fileId === fileId ? { ...m, column } : m
            )
          };
        } else {
          return {
            ...mapping,
            mappings: [...mapping.mappings, { fileId, column }]
          };
        }
      }
      return mapping;
    });

    onMappingsChange(updatedMappings);
  };

  const getAvailableColumns = (fileId: string) => {
    const worksheet = selectedWorksheets.find(ws => ws.fileId === fileId);
    return worksheet?.columns || [];
  };

  const getFileName = (fileId: string) => {
    const file = selectedFiles.find(f => f.id === fileId);
    return file ? file.name : fileId.replace('file-', 'File ');
  };

  const getMappedColumnsCount = () => {
    return columnMappings.reduce((sum, mapping) => sum + mapping.mappings.length, 0);
  };

  const getTotalColumnsCount = () => {
    return selectedWorksheets.reduce((sum, ws) => sum + ws.columns.length, 0);
  };

  const getMappingCompletionPercent = () => {
    const total = getTotalColumnsCount();
    const mapped = getMappedColumnsCount();
    return total > 0 ? Math.round((mapped / total) * 100) : 0;
  };

  const isPerfectMapping = () => {
    return getMappingCompletionPercent() === 100 && columnMappings.length > 0;
  };

  const canProceed = columnMappings.length > 0 && columnMappings.every(m => m.mappings.length > 0);

  return (
    <div className="p-8">
      <CardHeader className="text-center">
        <CardTitle className="text-2xl text-excel-primary">Column Mapping</CardTitle>
        <CardDescription className="text-lg">
          Map columns from different worksheets to create unified output
        </CardDescription>
      </CardHeader>

      <CardContent className="space-y-6">
        {/* Auto Map Section */}
        <Card className="bg-gradient-secondary/10 border-excel-secondary/20">
          <CardContent className="p-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-3">
                <Wand2 className="h-5 w-5 text-excel-secondary" />
                <div>
                  <p className="font-medium">Intelligent Auto-Mapping</p>
                  <p className="text-sm text-muted-foreground">
                    Automatically match columns with similar names
                  </p>
                </div>
              </div>
              <Button 
                onClick={autoMap}
                variant="secondary"
                className="bg-excel-secondary hover:bg-excel-secondary/90"
              >
                <Wand2 className="mr-2 h-4 w-4" />
                Auto Map
              </Button>
            </div>
          </CardContent>
        </Card>

        {/* Mapping Status */}
        <Card className={`transition-all duration-300 ${
          isPerfectMapping() 
            ? 'bg-green-50 border-green-200 dark:bg-green-900/20 dark:border-green-800' 
            : getMappingCompletionPercent() > 0 
              ? 'bg-orange-50 border-orange-200 dark:bg-orange-900/20 dark:border-orange-800'
              : 'bg-red-50 border-red-200 dark:bg-red-900/20 dark:border-red-800'
        }`}>
          <CardContent className="p-4">
            <div className="flex items-center justify-between">
              <div className="flex items-center space-x-3">
                {isPerfectMapping() ? (
                  <CheckCircle className="h-5 w-5 text-green-600 dark:text-green-400" />
                ) : (
                  <XCircle className="h-5 w-5 text-red-600 dark:text-red-400" />
                )}
                <span className="font-medium">
                  {getMappedColumnsCount()} of {getTotalColumnsCount()} columns mapped
                </span>
              </div>
              <div className="flex items-center space-x-2">
                <Badge 
                  variant={isPerfectMapping() ? "default" : "secondary"}
                  className={`transition-colors duration-300 ${
                    isPerfectMapping() 
                      ? 'bg-green-600 hover:bg-green-700 text-white dark:bg-green-500 dark:hover:bg-green-600' 
                      : getMappingCompletionPercent() > 0
                        ? 'bg-orange-500 hover:bg-orange-600 text-white'
                        : 'bg-red-500 hover:bg-red-600 text-white'
                  }`}
                >
                  {getMappingCompletionPercent()}% Complete
                </Badge>
                {isPerfectMapping() && (
                  <div className="flex items-center space-x-1 text-green-600 dark:text-green-400">
                    <CheckCircle className="h-4 w-4" />
                    <span className="text-sm font-medium">Perfect!</span>
                  </div>
                )}
              </div>
            </div>
          </CardContent>
        </Card>

        {/* Column Mappings */}
        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-lg font-medium text-excel-primary">Output Columns</h3>
            <div className="flex items-center space-x-2">
              <Input
                placeholder="New column name..."
                value={newColumnName}
                onChange={(e) => setNewColumnName(e.target.value)}
                className="w-48"
                onKeyPress={(e) => e.key === 'Enter' && addNewMapping()}
              />
              <Button 
                onClick={addNewMapping}
                size="sm"
                disabled={!newColumnName.trim()}
              >
                <Plus className="h-4 w-4" />
              </Button>
            </div>
          </div>

          {columnMappings.length === 0 ? (
            <Card className="border-dashed border-2 border-muted-foreground/25">
              <CardContent className="p-8 text-center">
                <AlertTriangle className="mx-auto h-12 w-12 text-muted-foreground mb-4" />
                <p className="text-lg font-medium mb-2">No column mappings defined</p>
                <p className="text-muted-foreground mb-4">
                  Use Auto Map or manually create column mappings to proceed
                </p>
                <Button onClick={autoMap} variant="outline">
                  <Wand2 className="mr-2 h-4 w-4" />
                  Auto Map Columns
                </Button>
              </CardContent>
            </Card>
          ) : (
            <div className="space-y-3">
              {columnMappings.map((mapping, mappingIndex) => (
                <Card key={mappingIndex} className="border-excel-primary/20">
                  <CardContent className="p-4">
                    <div className="flex items-center justify-between mb-3">
                      <Label className="text-base font-medium text-excel-primary">
                        {mapping.outputColumn}
                      </Label>
                      <Button
                        variant="ghost"
                        size="sm"
                        onClick={() => removeMapping(mappingIndex)}
                        className="text-destructive hover:text-destructive"
                      >
                        <X className="h-4 w-4" />
                      </Button>
                    </div>
                    
                    <div className="grid gap-3 md:grid-cols-2 lg:grid-cols-3">
                      {selectedWorksheets.map((worksheet) => {
                        const currentMapping = mapping.mappings.find(m => m.fileId === worksheet.fileId);
                        
                        return (
                          <div key={worksheet.fileId} className="space-y-2">
                            <Label className="text-sm text-muted-foreground">
                              {getFileName(worksheet.fileId)}
                            </Label>
                            <Select
                              value={currentMapping?.column || '__none__'}
                              onValueChange={(value) => 
                                updateMappingColumn(mappingIndex, worksheet.fileId, value === '__none__' ? '' : value)
                              }
                            >
                              <SelectTrigger>
                                <SelectValue placeholder="Select column..." />
                              </SelectTrigger>
                              <SelectContent className="bg-white z-50">
                                <SelectItem value="__none__">No mapping</SelectItem>
                                {getAvailableColumns(worksheet.fileId)
                                  .filter(column => column && typeof column === 'string' && column.trim().length > 0)
                                  .map((column) => (
                                    <SelectItem key={column} value={column}>
                                      {column}
                                    </SelectItem>
                                  ))}
                              </SelectContent>
                            </Select>
                          </div>
                        );
                      })}
                    </div>
                  </CardContent>
                </Card>
              ))}
            </div>
          )}
        </div>

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
            Combine & Save
          </Button>
        </div>
      </CardContent>
    </div>
  );
}