import React, { useState } from 'react';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Input } from '../ui/input';
import { Label } from '../ui/label';
import { Badge } from '../ui/badge';
import { Checkbox } from '../ui/checkbox';
import { Switch } from '../ui/switch';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '../ui/select';
import { ArrowLeft, Wand2, Plus, X, Link, AlertTriangle, CheckCircle, XCircle, Loader2 } from 'lucide-react';
import { WorksheetData, ColumnMapping as ColumnMappingType, ExcelFile } from '../ExcelCombiner';
import { cn } from '../../lib/utils';

interface ColumnMappingProps {
  selectedFiles: ExcelFile[];
  selectedWorksheets: WorksheetData[];
  columnMappings: ColumnMappingType[];
  onMappingsChange: (mappings: ColumnMappingType[]) => void;
  allowIncompleteMappings: boolean;
  onAllowIncompleteMappingsChange: (value: boolean) => void;
  allowDoubleMapping: boolean;
  onAllowDoubleMappingChange: (value: boolean) => void;
  onNext: () => void;
  onBack: () => void;
  isProcessing?: boolean;
}

export function ColumnMapping({
  selectedFiles,
  selectedWorksheets,
  columnMappings,
  onMappingsChange,
  allowIncompleteMappings,
  onAllowIncompleteMappingsChange,
  allowDoubleMapping,
  onAllowDoubleMappingChange,
  onNext,
  onBack,
  isProcessing = false
}: ColumnMappingProps) {
  const [newColumnName, setNewColumnName] = useState('');
  const [hideMappedColumns, setHideMappedColumns] = useState(false);

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

  const isMappingFullyIncomplete = (mapping: ColumnMappingType) => {
    return selectedWorksheets.every((worksheet) => {
      const fileMapping = mapping.mappings.find(m => m.fileId === worksheet.fileId);
      return !fileMapping?.column?.trim();
    });
  };

  const deleteFullyIncompleteMappings = () => {
    onMappingsChange(columnMappings.filter(mapping => !isMappingFullyIncomplete(mapping)));
  };

  const updateMappingColumn = (mappingIndex: number, fileId: string, column: string) => {
    const updatedMappings = columnMappings.map((mapping, index) => {
      if (index === mappingIndex) {
        if (!column || column.trim().length === 0) {
          // Remove the mapping entirely when "No mapping" is selected
          return {
            ...mapping,
            mappings: mapping.mappings.filter(m => m.fileId !== fileId)
          };
        }
        
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

    if (!allowDoubleMapping && column.trim()) {
      const normalizedMappings = updatedMappings.map((mapping, index) => {
        if (index === mappingIndex) {
          return mapping;
        }

        return {
          ...mapping,
          mappings: mapping.mappings.filter(
            (fileMapping) => !(fileMapping.fileId === fileId && fileMapping.column === column)
          ),
        };
      });

      onMappingsChange(normalizedMappings);
      return;
    }

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
    // Count only output columns that are fully mapped across every selected worksheet
    return columnMappings.filter(mapping => isMappingComplete(mapping)).length;
  };

  const getTotalColumnsCount = () => {
    // Total number of output columns
    return columnMappings.length;
  };

  const getMappingCompletionPercent = () => {
    const total = getTotalColumnsCount();
    const mapped = getMappedColumnsCount();
    return total > 0 ? Math.round((mapped / total) * 100) : 0;
  };

  const isPerfectMapping = () => {
    return getMappingCompletionPercent() === 100 && columnMappings.length > 0;
  };

  const hasAnyMappedColumn = (mapping: ColumnMappingType) => {
    return mapping.mappings.some(fileMapping => !!fileMapping.column?.trim());
  };

  const isMappingComplete = (mapping: ColumnMappingType) => {
    if (selectedWorksheets.length === 0) {
      return false;
    }

    return selectedWorksheets.every((worksheet) => {
      const fileMapping = mapping.mappings.find(m => m.fileId === worksheet.fileId);
      return !!fileMapping?.column?.trim();
    });
  };

  const getMissingMappings = (mapping: ColumnMappingType) => {
    return selectedWorksheets.filter((worksheet) => {
      const fileMapping = mapping.mappings.find(m => m.fileId === worksheet.fileId);
      return !fileMapping?.column?.trim();
    });
  };

  const incompleteMappings = columnMappings.filter(mapping => !isMappingComplete(mapping));
  const hasIncompleteMappings = incompleteMappings.length > 0;
  const fullyIncompleteMappings = columnMappings.filter(isMappingFullyIncomplete);
  const visibleMappings = hideMappedColumns
    ? columnMappings.filter(mapping => !isMappingComplete(mapping))
    : columnMappings;

  const canProceed = columnMappings.length > 0 && columnMappings.every(mapping =>
    allowIncompleteMappings ? hasAnyMappedColumn(mapping) : isMappingComplete(mapping)
  );

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
            <div className="flex items-center justify-between gap-4">
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

            <div className="mt-4 flex flex-col gap-2 rounded-lg border border-excel-secondary/20 bg-background/80 p-3 sm:flex-row sm:items-center sm:justify-between">
              <div>
                <Label htmlFor="allow-double-mapping" className="text-sm font-medium text-excel-primary">
                  Allow double mapping
                </Label>
                <p className="text-xs text-muted-foreground">
                  Off by default. When off, choosing a source column removes that same file-column from other output columns.
                </p>
              </div>
              <div className="flex items-center gap-3">
                <span className={cn('text-sm font-medium', allowDoubleMapping ? 'text-amber-600' : 'text-muted-foreground')}>
                  {allowDoubleMapping ? 'On' : 'Off'}
                </span>
                <Switch
                  id="allow-double-mapping"
                  checked={allowDoubleMapping}
                  onCheckedChange={onAllowDoubleMappingChange}
                />
              </div>
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

        {hasIncompleteMappings && (
          <Card className="border-red-300 bg-red-50 dark:border-red-800 dark:bg-red-950/20">
            <CardContent className="p-4">
              <div className="flex items-start gap-3">
                <AlertTriangle className="mt-0.5 h-5 w-5 text-red-600 dark:text-red-400" />
                <div className="flex-1 space-y-3">
                  <div>
                    <p className="font-medium text-red-700 dark:text-red-400">
                      {incompleteMappings.length} output {incompleteMappings.length === 1 ? 'column is' : 'columns are'} incomplete
                    </p>
                    <p className="text-sm text-red-600 dark:text-red-300">
                      Incomplete mappings will leave blank values for files that are not mapped. Fix them now, or explicitly accept them to continue.
                    </p>
                  </div>

                  <div className="flex items-start gap-3 rounded-md border border-red-200 bg-white/80 p-3 dark:border-red-800 dark:bg-red-950/30">
                    <Checkbox
                      id="accept-incomplete-mappings"
                      checked={allowIncompleteMappings}
                      onCheckedChange={(checked) => onAllowIncompleteMappingsChange(checked === true)}
                      className="mt-0.5 border-red-500 data-[state=checked]:bg-red-600 data-[state=checked]:text-white"
                    />
                    <Label
                      htmlFor="accept-incomplete-mappings"
                      className="cursor-pointer text-sm leading-5 text-red-700 dark:text-red-300"
                    >
                      I understand that incomplete mappings will be kept with blanks for missing files, and I want to continue anyway.
                    </Label>
                  </div>
                </div>
              </div>
            </CardContent>
          </Card>
        )}

        {/* Column Mappings */}
        <div className="space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-lg font-medium text-excel-primary">Output Columns</h3>
            <div className="flex flex-wrap items-center justify-end gap-2">
              <div className="flex items-center gap-2 rounded-md border bg-background px-3 py-2">
                <Label htmlFor="hide-mapped-columns" className="text-sm font-medium text-excel-primary">
                  Hide mapped columns
                </Label>
                <Switch
                  id="hide-mapped-columns"
                  checked={hideMappedColumns}
                  onCheckedChange={setHideMappedColumns}
                />
              </div>
              <Button
                variant="outline"
                size="sm"
                onClick={deleteFullyIncompleteMappings}
                disabled={fullyIncompleteMappings.length === 0}
              >
                Delete fully incomplete ({fullyIncompleteMappings.length})
              </Button>
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

          {hideMappedColumns && (
            <p className="text-sm text-muted-foreground">
              Showing {visibleMappings.length} incomplete column{visibleMappings.length === 1 ? '' : 's'}. Fully mapped columns are hidden automatically.
            </p>
          )}

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
              {visibleMappings.length === 0 ? (
                <Card className="border-dashed border-2 border-muted-foreground/25">
                  <CardContent className="p-8 text-center">
                    <CheckCircle className="mx-auto mb-4 h-12 w-12 text-excel-accent-green" />
                    <p className="text-lg font-medium mb-2">No columns to show</p>
                    <p className="text-muted-foreground">
                      All visible output columns are fully mapped.
                    </p>
                  </CardContent>
                </Card>
              ) : visibleMappings.map((mapping) => {
                const mappingIndex = columnMappings.indexOf(mapping);
                return (
                <Card
                  key={mappingIndex}
                  className={cn(
                    'transition-all duration-200',
                    isMappingComplete(mapping)
                      ? 'border-excel-primary/20'
                      : 'border-red-300 bg-red-50/80 shadow-sm dark:border-red-800 dark:bg-red-950/20'
                  )}
                >
                  <CardContent className="p-4">
                    <div className="flex items-center justify-between mb-3">
                      <div className="space-y-1">
                        <Label className={cn(
                          'text-base font-medium',
                          isMappingComplete(mapping) ? 'text-excel-primary' : 'text-red-700 dark:text-red-400'
                        )}>
                          {mapping.outputColumn}
                        </Label>
                        {!isMappingComplete(mapping) && (
                          <div className="flex items-center gap-2 text-sm text-red-600 dark:text-red-400">
                            <AlertTriangle className="h-4 w-4" />
                            <span>
                              Unmapped for {getMissingMappings(mapping).map((worksheet) => getFileName(worksheet.fileId)).join(', ')}
                            </span>
                          </div>
                        )}
                      </div>
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
                        const hasMissingMapping = !currentMapping?.column || !currentMapping.column.trim();
                        
                        return (
                          <div key={worksheet.fileId} className="flex flex-col gap-2">
                            <div className="min-h-[3rem] space-y-1">
                              <Label className="text-sm text-muted-foreground">
                                {getFileName(worksheet.fileId)}
                              </Label>
                              <p className={cn(
                                'text-xs font-medium min-h-[1rem]',
                                hasMissingMapping
                                  ? 'text-red-600 dark:text-red-400'
                                  : 'invisible'
                              )}>
                                No mapping selected
                              </p>
                            </div>
                            <Select
                              value={currentMapping?.column || '__none__'}
                              onValueChange={(value) => 
                                updateMappingColumn(mappingIndex, worksheet.fileId, value === '__none__' ? '' : value)
                              }
                            >
                              <SelectTrigger
                                className={cn(
                                  hasMissingMapping &&
                                    'border-red-400 text-red-600 focus:ring-red-500 dark:border-red-700 dark:text-red-400'
                                )}
                              >
                                <SelectValue placeholder="Select column..." />
                              </SelectTrigger>
                              <SelectContent className="bg-white z-50">
                                <SelectItem
                                  value="__none__"
                                  className="text-red-600 focus:text-red-700 dark:text-red-400 dark:focus:text-red-300"
                                >
                                  No mapping
                                </SelectItem>
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
              )})}
            </div>
          )}
        </div>

        {/* Navigation */}
        <div className="flex justify-between pt-6">
          <Button variant="outline" onClick={onBack} disabled={isProcessing}>
            <ArrowLeft className="mr-2 h-4 w-4" />
            Back to Preview
          </Button>
          
          <Button
            onClick={onNext}
            disabled={!canProceed || isProcessing}
            className="bg-gradient-primary hover:opacity-90 px-8"
          >
            {isProcessing ? (
              <>
                <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                Processing Files...
              </>
            ) : (
              'Next: Rearrange Columns'
            )}
          </Button>
        </div>
      </CardContent>
    </div>
  );
}