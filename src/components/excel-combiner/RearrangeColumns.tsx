import React from 'react';
import { ArrowLeft, CheckCircle, GripVertical, Loader2, MousePointerClick } from 'lucide-react';
import { Button } from '../ui/button';
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from '../ui/card';
import { Input } from '../ui/input';
import { Badge } from '../ui/badge';
import { Label } from '../ui/label';
import { ColumnMapping, WorksheetData } from '../ExcelCombiner';
import { cn } from '../../lib/utils';

interface RearrangeColumnsProps {
  columnMappings: ColumnMapping[];
  selectedWorksheets: WorksheetData[];
  onMappingsChange: (mappings: ColumnMapping[]) => void;
  onNext: () => void;
  onBack: () => void;
  isProcessing?: boolean;
}

export function RearrangeColumns({
  columnMappings,
  selectedWorksheets,
  onMappingsChange,
  onNext,
  onBack,
  isProcessing = false,
}: RearrangeColumnsProps) {
  const rowRefs = React.useRef<Array<HTMLDivElement | null>>([]);
  const autoScrollFrameRef = React.useRef<number | null>(null);
  const [selectedIndexes, setSelectedIndexes] = React.useState<number[]>([]);
  const [anchorIndex, setAnchorIndex] = React.useState<number | null>(null);
  const [dragState, setDragState] = React.useState<{
    indexes: number[];
    pointerY: number;
    dropIndex: number;
  } | null>(null);

  React.useEffect(() => {
    setSelectedIndexes((prev) => prev.filter((index) => index < columnMappings.length));
    setAnchorIndex((prev) => (prev !== null && prev < columnMappings.length ? prev : null));
  }, [columnMappings.length]);

  React.useEffect(() => {
    return () => {
      if (autoScrollFrameRef.current !== null) {
        cancelAnimationFrame(autoScrollFrameRef.current);
      }
    };
  }, []);

  const mappedCount = (mapping: ColumnMapping) => {
    return selectedWorksheets.filter((worksheet) =>
      mapping.mappings.some((fileMapping) => fileMapping.fileId === worksheet.fileId && !!fileMapping.column?.trim())
    ).length;
  };

  const handleNameChange = (index: number, value: string) => {
    onMappingsChange(
      columnMappings.map((mapping, mappingIndex) =>
        mappingIndex === index ? { ...mapping, outputColumn: value } : mapping
      )
    );
  };

  const handleSelect = (index: number, event: React.MouseEvent<HTMLDivElement>) => {
    if (event.shiftKey && anchorIndex !== null) {
      const start = Math.min(anchorIndex, index);
      const end = Math.max(anchorIndex, index);
      setSelectedIndexes(Array.from({ length: end - start + 1 }, (_, offset) => start + offset));
      return;
    }

    if (event.ctrlKey || event.metaKey) {
      setSelectedIndexes((prev) => {
        if (prev.includes(index)) {
          return prev.filter((item) => item !== index);
        }

        return [...prev, index].sort((a, b) => a - b);
      });
      setAnchorIndex(index);
      return;
    }

    setSelectedIndexes([index]);
    setAnchorIndex(index);
  };

  const moveSelectedMappings = React.useCallback((targetIndex: number, sourceIndexes: number[]) => {
    if (sourceIndexes.length === 0) {
      return;
    }

    const selectedSet = new Set(sourceIndexes);
    const movedItems = sourceIndexes.map((index) => columnMappings[index]);
    const remainingItems = columnMappings.filter((_, index) => !selectedSet.has(index));
    const insertionIndex = Math.max(0, targetIndex - sourceIndexes.filter((index) => index < targetIndex).length);
    const reorderedMappings = [
      ...remainingItems.slice(0, insertionIndex),
      ...movedItems,
      ...remainingItems.slice(insertionIndex),
    ];

    onMappingsChange(reorderedMappings);
    const nextSelection = movedItems.map((_, offset) => insertionIndex + offset);
    setSelectedIndexes(nextSelection);
    setAnchorIndex(nextSelection[0] ?? null);
  }, [columnMappings, onMappingsChange]);

  const getDropIndexFromPointer = React.useCallback((pointerY: number) => {
    for (let index = 0; index < columnMappings.length; index++) {
      const row = rowRefs.current[index];
      if (!row) {
        continue;
      }

      const rect = row.getBoundingClientRect();
      const midpoint = rect.top + rect.height / 2;
      if (pointerY < midpoint) {
        return index;
      }
    }

    return columnMappings.length;
  }, [columnMappings.length]);

  const stopDragging = React.useCallback(() => {
    setDragState(null);
    document.body.style.userSelect = '';
    document.body.style.cursor = '';
    if (autoScrollFrameRef.current !== null) {
      cancelAnimationFrame(autoScrollFrameRef.current);
      autoScrollFrameRef.current = null;
    }
  }, []);

  const handleDragStart = (index: number, event: React.MouseEvent<HTMLButtonElement>) => {
    event.preventDefault();
    event.stopPropagation();

    const nextDraggedIndexes = (selectedIndexes.includes(index) ? selectedIndexes : [index]).sort((a, b) => a - b);
    if (!selectedIndexes.includes(index)) {
      setSelectedIndexes([index]);
      setAnchorIndex(index);
    }

    document.body.style.userSelect = 'none';
    document.body.style.cursor = 'grabbing';

    setDragState({
      indexes: nextDraggedIndexes,
      pointerY: event.clientY,
      dropIndex: getDropIndexFromPointer(event.clientY),
    });
  };

  React.useEffect(() => {
    if (!dragState) {
      return;
    }

    const handleMouseMove = (event: MouseEvent) => {
      setDragState((prev) => prev ? {
        ...prev,
        pointerY: event.clientY,
        dropIndex: getDropIndexFromPointer(event.clientY),
      } : null);
    };

    const handleMouseUp = () => {
      moveSelectedMappings(dragState.dropIndex, dragState.indexes);
      stopDragging();
    };

    const autoScroll = () => {
      setDragState((prev) => {
        if (!prev) {
          return prev;
        }

        const threshold = 120;
        const maxSpeed = 18;
        let scrollDelta = 0;

        if (prev.pointerY < threshold) {
          scrollDelta = -Math.min(maxSpeed, Math.ceil((threshold - prev.pointerY) / 10));
        } else if (window.innerHeight - prev.pointerY < threshold) {
          scrollDelta = Math.min(maxSpeed, Math.ceil((threshold - (window.innerHeight - prev.pointerY)) / 10));
        }

        if (scrollDelta !== 0) {
          window.scrollBy(0, scrollDelta);
        }

        autoScrollFrameRef.current = requestAnimationFrame(autoScroll);
        return {
          ...prev,
          dropIndex: getDropIndexFromPointer(prev.pointerY),
        };
      });
    };

    window.addEventListener('mousemove', handleMouseMove);
    window.addEventListener('mouseup', handleMouseUp);
    autoScrollFrameRef.current = requestAnimationFrame(autoScroll);

    return () => {
      window.removeEventListener('mousemove', handleMouseMove);
      window.removeEventListener('mouseup', handleMouseUp);
      if (autoScrollFrameRef.current !== null) {
        cancelAnimationFrame(autoScrollFrameRef.current);
        autoScrollFrameRef.current = null;
      }
    };
  }, [dragState, getDropIndexFromPointer, moveSelectedMappings, stopDragging]);

  const canProceed = columnMappings.length > 0 && columnMappings.every((mapping) => mapping.outputColumn.trim().length > 0);

  return (
    <div className="p-8">
      <CardHeader className="text-center">
        <CardTitle className="text-2xl text-excel-primary">Rearrange Columns</CardTitle>
        <CardDescription className="text-lg">
          Reorder the final output columns and confirm their names before generating the CSV
        </CardDescription>
      </CardHeader>

      <CardContent className="space-y-4">
        <Card className="bg-gradient-secondary/10 border-excel-secondary/20">
          <CardContent className="p-4">
            <div className="flex items-start gap-3">
              <MousePointerClick className="mt-0.5 h-5 w-5 text-excel-secondary" />
              <div className="space-y-1 text-sm text-muted-foreground">
                <p className="font-medium text-foreground">How to use</p>
                <p>Click to select one column. Use Ctrl/Cmd-click to multi-select. Use Shift-click to select a range. Drag any selected row to move the whole selection together.</p>
              </div>
            </div>
          </CardContent>
        </Card>

        <div className="space-y-2 relative">
          {dragState && dragState.dropIndex === 0 && (
            <div className="h-0 rounded-full border-2 border-excel-secondary shadow-sm shadow-excel-secondary/30" />
          )}
          {columnMappings.map((mapping, index) => {
            const isSelected = selectedIndexes.includes(index);
            const isDragged = !!dragState?.indexes.includes(index);
            const mappedFiles = mappedCount(mapping);

            return (
              <React.Fragment key={`${mapping.outputColumn}-${index}`}>
              <Card
                ref={(element) => {
                  rowRefs.current[index] = element;
                }}
                onClick={(event) => handleSelect(index, event)}
                className={cn(
                  'cursor-pointer border transition-all duration-150 select-none',
                  isSelected && 'border-excel-primary bg-excel-primary/5 shadow-sm',
                  isDragged && 'opacity-35',
                  !isSelected && 'border-excel-primary/20'
                )}
              >
                <CardContent className="p-3">
                  <div className="grid gap-3 md:grid-cols-[180px_minmax(220px,0.8fr)_minmax(360px,1.4fr)] md:items-center">
                    <div className="flex items-center gap-3">
                      <button
                        type="button"
                        onMouseDown={(event) => handleDragStart(index, event)}
                        className="rounded-md border bg-muted/40 p-2 text-muted-foreground transition hover:bg-muted/70 active:cursor-grabbing cursor-grab"
                        aria-label={`Drag column ${mapping.outputColumn}`}
                      >
                        <GripVertical className="h-4 w-4" />
                      </button>
                      <div className="space-y-0.5">
                        <p className="text-sm font-medium text-muted-foreground">Position {index + 1}</p>
                        <div className="flex flex-wrap items-center gap-2">
                          <Badge variant="secondary">{mappedFiles}/{selectedWorksheets.length} files mapped</Badge>
                          {isSelected && <Badge className="bg-excel-primary text-white">Selected</Badge>}
                        </div>
                      </div>
                    </div>

                    <div className="space-y-1.5">
                      <Label htmlFor={`output-column-${index}`}>Column name</Label>
                      <Input
                        id={`output-column-${index}`}
                        value={mapping.outputColumn}
                        onChange={(event) => handleNameChange(index, event.target.value)}
                        onClick={(event) => event.stopPropagation()}
                        placeholder="Enter output column name"
                      />
                    </div>

                    <div className="space-y-1.5">
                      <Label>Current source mappings</Label>
                      <div className="rounded-md border bg-muted/30 px-3 py-2 text-sm leading-5 text-muted-foreground break-words min-h-[64px]">
                        {mapping.mappings.length > 0
                          ? mapping.mappings.map((fileMapping) => fileMapping.column).join(', ')
                          : 'No mapped source columns'}
                      </div>
                    </div>
                  </div>
                </CardContent>
              </Card>
              {dragState && dragState.dropIndex === index + 1 && (
                <div className="h-0 rounded-full border-2 border-excel-secondary shadow-sm shadow-excel-secondary/30" />
              )}
              </React.Fragment>
            );
          })}
        </div>

        {dragState && (
          <div className="pointer-events-none fixed bottom-6 right-6 z-50">
            <div className="relative h-20 w-[320px]">
              {dragState.indexes.slice(0, 3).map((index, layerIndex) => {
                const mapping = columnMappings[index];
                return (
                  <div
                    key={`${mapping.outputColumn}-${layerIndex}`}
                    className={cn(
                      'absolute inset-0 rounded-xl border border-excel-primary/30 bg-background/95 shadow-lg',
                      layerIndex === 0 && 'translate-x-0 translate-y-0 opacity-100',
                      layerIndex === 1 && 'translate-x-2 translate-y-2 opacity-80',
                      layerIndex === 2 && 'translate-x-4 translate-y-4 opacity-60'
                    )}
                  >
                    <div className="flex h-full items-center justify-between px-4 py-3">
                      <div className="min-w-0">
                        <p className="truncate text-sm font-semibold text-foreground">{mapping.outputColumn}</p>
                        <p className="truncate text-xs text-muted-foreground">
                          {mappedCount(mapping)}/{selectedWorksheets.length} files mapped
                        </p>
                      </div>
                      {layerIndex === 0 && (
                        <Badge className="bg-excel-primary text-white">
                          {dragState.indexes.length} selected
                        </Badge>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        )}

        <div className="flex justify-between pt-4">
          <Button variant="outline" onClick={onBack} disabled={isProcessing}>
            <ArrowLeft className="mr-2 h-4 w-4" />
            Back to Mapping
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
              <>
                <CheckCircle className="mr-2 h-4 w-4" />
                Generate Combined CSV
              </>
            )}
          </Button>
        </div>
      </CardContent>
    </div>
  );
}
