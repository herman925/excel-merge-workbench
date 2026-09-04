import React, { useRef, useState } from 'react';
import { Button } from '../ui/button';
import {
  Dialog, DialogContent, DialogHeader, DialogTitle, DialogDescription,
  DialogFooter, DialogClose,
} from '../ui/dialog';
import { Badge } from '../ui/badge';
import { Input } from '../ui/input';
import { Label } from '../ui/label';
import { useToast } from '../../hooks/use-toast';
import {
  Download, Upload, Save, Trash2, FileSpreadsheet, CheckCircle2, ChevronRight,
} from 'lucide-react';
import type { ExcelFile, WorksheetData, ColumnMapping } from '../ExcelCombiner';
import {
  MergeConfig, ConfigPreset, buildMergeConfig, parseMergeConfig, downloadConfig,
  configFilename, defaultPresetName, listPresets, savePreset, removePreset, buildExcelFile, readWorksheetColumns,
} from '../../lib/merge-config';

export interface PendingItem {
  step: string;
  message: string;
}

export interface ImportResult {
  files: ExcelFile[];
  worksheets: WorksheetData[];
  columnMappings: ColumnMapping[];
  keyColumn: string;
  allowIncompleteMappings: boolean;
  allowDoubleMapping: boolean;
  pending: PendingItem[];
}

interface ConfigManagerProps {
  files: ExcelFile[];
  worksheets: WorksheetData[];
  columnMappings: ColumnMapping[];
  keyColumn: string;
  allowIncompleteMappings: boolean;
  allowDoubleMapping: boolean;
  onApplyImport: (result: ImportResult) => void;
}

type ModalMode = 'none' | 'presets' | 'import';

export function ConfigManager({
  files, worksheets, columnMappings, keyColumn, allowIncompleteMappings, allowDoubleMapping,
  onApplyImport,
}: ConfigManagerProps) {
  const { toast } = useToast();
  const [modalMode, setModalMode] = useState<ModalMode>('none');
  const [presets, setPresets] = useState<ConfigPreset[]>(listPresets());

  // import re-point state
  const [importConfig, setImportConfig] = useState<MergeConfig | null>(null);
  const [importStep, setImportStep] = useState(0);
  const [pickedFiles, setPickedFiles] = useState<Record<number, File>>({});
  const [skipped, setSkipped] = useState<Set<number>>(new Set());
  const [processing, setProcessing] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const pointFileInputRef = useRef<HTMLInputElement>(null);

  const handleExport = () => {
    const cfg = buildMergeConfig(files, worksheets, columnMappings, keyColumn, allowIncompleteMappings, allowDoubleMapping);
    if (cfg.files.length === 0) {
      toast({ title: 'Nothing to export', description: 'Select files and configure a merge first.', variant: 'destructive' });
      return;
    }
    downloadConfig(cfg, configFilename(cfg));
    const name = defaultPresetName(cfg);
    savePreset(name, cfg);
    setPresets(listPresets());
    toast({ title: 'Config exported', description: `${configFilename(cfg)} downloaded and saved as a preset.` });
  };

  const handleSavePreset = () => {
    const cfg = buildMergeConfig(files, worksheets, columnMappings, keyColumn, allowIncompleteMappings, allowDoubleMapping);
    if (cfg.files.length === 0) {
      toast({ title: 'Nothing to save', description: 'Select files and configure a merge first.', variant: 'destructive' });
      return;
    }
    savePreset(defaultPresetName(cfg), cfg);
    setPresets(listPresets());
    toast({ title: 'Preset saved', description: `Saved "${defaultPresetName(cfg)}".` });
  };

  const handleImportFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file) return;
    if (file.size > 5_000_000) {
      toast({ title: 'Config file too large', description: 'Merge configs are small JSON files — this one is over 5 MB.', variant: 'destructive' });
      return;
    }
    try {
      const text = await file.text();
      const cfg = parseMergeConfig(text);
      startImport(cfg);
    } catch (error) {
      toast({ title: 'Import failed', description: error instanceof Error ? error.message : 'Could not read config.', variant: 'destructive' });
    }
  };

  const startImport = (cfg: MergeConfig) => {
    setImportConfig(cfg);
    setImportStep(0);
    setPickedFiles({});
    setSkipped(new Set());
    setModalMode('import');
  };

  const handlePickFile = async (index: number, file: File) => {
    setPickedFiles((prev) => ({ ...prev, [index]: file }));
    // advance to next unresolved slot
    advancePastResolved(index);
  };

  const handlePointFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    event.target.value = '';
    if (!file || !importConfig) return;
    if (!(file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
      toast({ title: 'Invalid file type', description: 'Please choose an .xlsx or .xls file.', variant: 'destructive' });
      return;
    }
    handlePickFile(importStep, file);
  };

  const advancePastResolved = (from: number) => {
    if (!importConfig) return;
    let next = from + 1;
    while (next < importConfig.files.length && (pickedFiles[next] || skipped.has(next))) {
      next++;
    }
    if (next < importConfig.files.length) {
      setImportStep(next);
    } else {
      setImportStep(importConfig.files.length);
    }
  };

  const handleSkip = () => {
    if (!importConfig) return;
    setSkipped((prev) => new Set(prev).add(importStep));
    advancePastResolved(importStep);
  };

  const applyImport = async () => {
    if (!importConfig) return;
    setProcessing(true);
    try {
      const pending: PendingItem[] = [];
      const importedFiles: ExcelFile[] = [];
      const fileIndexToId = new Map<number, string>();

      // Build resolved ExcelFile objects in config order
      for (let i = 0; i < importConfig.files.length; i++) {
        const file = pickedFiles[i];
        if (!file) {
          pending.push({ step: 'file-selection', message: importConfig.files[i].name });
          continue;
        }
        const id = `import-${i}`;
        fileIndexToId.set(i, id);
        importedFiles.push(await buildExcelFile(file, id));
      }

      // Ensure each imported file has an id in columnMappings' fileId space
      const newWorksheets: WorksheetData[] = [];
      for (const cw of importConfig.worksheets) {
        const id = fileIndexToId.get(cw.fileIndex);
        if (!id) continue;
        const ef = importedFiles.find((f) => f.id === id);
        if (!ef) continue;
        if (!ef.worksheets.includes(cw.worksheetName)) {
          pending.push({ step: 'worksheet-selection', message: `${ef.name} → ${cw.worksheetName}` });
          continue;
        }
        const columns = await readWorksheetColumns(ef.file, cw.worksheetName, cw.headerRow);
        newWorksheets.push({
          fileId: id,
          worksheetName: cw.worksheetName,
          headerRow: cw.headerRow,
          columns,
          ...(cw.keyColumn && columns.includes(cw.keyColumn) ? { keyColumn: cw.keyColumn } : {}),
        });
      }

      const worksheetMap = new Map(newWorksheets.map((w) => [w.fileId, w]));
      const newMappings: ColumnMapping[] = importConfig.columnMappings.map((m) => {
        const resolved = m.mappings
          .map((fm) => {
            const id = fileIndexToId.get(fm.fileIndex);
            const ws = id ? worksheetMap.get(id) : undefined;
            if (id && ws && ws.columns.includes(fm.column)) {
              return { fileId: id, column: fm.column };
            }
            return null;
          })
          .filter((fm): fm is { fileId: string; column: string } => !!fm);
        if (resolved.length < m.mappings.length) {
          pending.push({ step: 'column-mapping', message: `Mapping "${m.outputColumn}" lost a column` });
        }
        return { outputColumn: m.outputColumn, mappings: resolved };
      }).filter((m) => m.mappings.length > 0 || importConfig.allowIncompleteMappings);

      // keyColumn survives if it still exists in at least one worksheet
      let newKeyColumn = importConfig.keyColumn || '';
      if (newKeyColumn && !newWorksheets.some((w) => w.columns.includes(newKeyColumn))) {
        newKeyColumn = '';
      }

      onApplyImport({
        files: importedFiles,
        worksheets: newWorksheets,
        columnMappings: newMappings,
        keyColumn: newKeyColumn,
        allowIncompleteMappings: importConfig.allowIncompleteMappings,
        allowDoubleMapping: importConfig.allowDoubleMapping,
        pending,
      });

      setModalMode('none');
      setImportConfig(null);
      if (pending.length > 0) {
        toast({
          title: 'Import applied with gaps',
          description: `${pending.length} item(s) unresolved — marked pending for you to fix in the steps.`,
        });
      } else {
        toast({ title: 'Config imported', description: 'Merge setup restored. Review the steps and merge.' });
      }
    } catch (error) {
      toast({ title: 'Import failed', description: error instanceof Error ? error.message : 'Could not apply config.', variant: 'destructive' });
    } finally {
      setProcessing(false);
    }
  };

  const handleDeletePreset = (id: string) => {
    removePreset(id);
    setPresets(listPresets());
  };

  const handleLoadPreset = (preset: ConfigPreset) => {
    try {
      startImport(sanitizeMergeConfig(preset.config));
    } catch {
      toast({ title: 'Preset is corrupt', description: 'This saved preset is missing required sections and cannot be loaded.', variant: 'destructive' });
    }
  };

  const totalSlots = importConfig?.files.length ?? 0;
  const currentFile = importConfig?.files[importStep];

  return (
    <>
      <div className="flex items-center gap-2">
        <Button variant="outline" size="sm" onClick={handleExport} title="Export the merge config as a JSON file">
          <Download className="mr-1 h-4 w-4" /> Export
        </Button>
        <Button variant="outline" size="sm" onClick={() => fileInputRef.current?.click()} title="Import a merge config">
          <Upload className="mr-1 h-4 w-4" /> Import
        </Button>
        <Button variant="outline" size="sm" onClick={() => { setPresets(listPresets()); setModalMode('presets'); }} title="Saved presets">
          <Save className="mr-1 h-4 w-4" /> Presets
        </Button>
        <input ref={fileInputRef} type="file" accept="application/json,.json" onChange={handleImportFile} className="hidden" />
      </div>

      {/* Presets dialog */}
      <Dialog open={modalMode === 'presets'} onOpenChange={(o) => !o && setModalMode('none')}>
        <DialogContent className="max-w-md">
          <DialogHeader>
            <DialogTitle>Saved Merge Presets</DialogTitle>
            <DialogDescription>Re-load a saved merge setup without re-selecting the file.</DialogDescription>
          </DialogHeader>
          <div className="space-y-2 max-h-80 overflow-y-auto pr-2">
            {presets.length === 0 && (
              <p className="text-sm text-muted-foreground">No presets yet. Export a config to create one.</p>
            )}
            {presets.map((p) => (
              <div key={p.id} className="flex items-center justify-between gap-2 rounded-md border p-3">
                <button className="flex-1 text-left" onClick={() => handleLoadPreset(p)}>
                  <div className="flex items-center gap-2">
                    <FileSpreadsheet className="h-4 w-4 text-excel-secondary" />
                    <span className="font-medium text-sm">{p.name}</span>
                    <Badge variant="secondary" className="text-xs">{p.config.files.length} files</Badge>
                  </div>
                  <div className="text-xs text-muted-foreground mt-1">
                    {new Date(p.savedAt).toLocaleString()}
                  </div>
                </button>
                <Button variant="ghost" size="icon" onClick={() => handleDeletePreset(p.id)} title="Delete preset">
                  <Trash2 className="h-4 w-4 text-destructive" />
                </Button>
              </div>
            ))}
          </div>
          <DialogFooter>
            <DialogClose asChild>
              <Button variant="outline">Close</Button>
            </DialogClose>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      {/* Import re-point dialog */}
      <Dialog open={modalMode === 'import'} onOpenChange={(o) => { if (!o) setModalMode('none'); }}>
        <DialogContent className="max-w-lg">
          <DialogHeader>
            <DialogTitle>Restore Merge Setup</DialogTitle>
            <DialogDescription>
              Re-point each file from disk. Skip any you don't want; unresolved items are marked pending.
            </DialogDescription>
          </DialogHeader>

          {importConfig && currentFile && (
            <div className="space-y-4">
              <input ref={pointFileInputRef} type="file" accept=".xlsx,.xls" onChange={handlePointFile} className="hidden" />
              {/* progress */}
              <div className="flex items-center gap-1.5">
                {importConfig.files.map((f, i) => {
                  const solved = !!pickedFiles[i];
                  const isSkip = skipped.has(i);
                  const isCurrent = i === importStep;
                  let cls = 'bg-muted text-muted-foreground';
                  if (solved) cls = 'bg-excel-accent-green text-white';
                  else if (isSkip) cls = 'bg-orange-300 text-orange-900';
                  if (isCurrent) cls += ' ring-2 ring-excel-primary';
                  return (
                    <div key={i} title={f.name}
                      className={`flex h-7 w-7 items-center justify-center rounded-full text-xs ${cls}`}>
                      {solved ? <CheckCircle2 className="h-4 w-4" /> : isSkip ? 'S' : i + 1}
                    </div>
                  );
                })}
              </div>

              <div className="rounded-md border p-4 space-y-3">
                <p className="text-sm font-medium">File {importStep + 1} of {totalSlots}: <span className="text-excel-primary">{currentFile.name}</span></p>
                {pickedFiles[importStep] ? (
                  <div className="flex items-center gap-2 text-sm text-excel-accent-green">
                    <CheckCircle2 className="h-4 w-4" /> {pickedFiles[importStep].name}
                  </div>
                ) : (
                  <div className="flex items-center gap-2">
                    <Button variant="outline" size="sm" onClick={() => pointFileInputRef.current?.click()}>
                      <Upload className="mr-1 h-4 w-4" /> Select {currentFile.name}
                    </Button>
                    <Button variant="ghost" size="sm" onClick={handleSkip} className="text-muted-foreground">
                      Skip
                    </Button>
                  </div>
                )}
              </div>
            </div>
          )}

          <DialogFooter className="justify-between">
            <Button variant="ghost" onClick={() => setModalMode('none')}>Cancel</Button>
            <div className="flex items-center gap-2">
              {importStep > 0 && importStep < totalSlots && (
                <Button variant="outline" onClick={() => setImportStep(importStep - 1)}>Back</Button>
              )}
              {importStep >= totalSlots ? (
                <Button onClick={applyImport} disabled={processing} className="bg-excel-accent-green hover:opacity-90">
                  {processing ? 'Applying...' : 'Apply Config'}
                </Button>
              ) : (
                <Button variant="outline" onClick={() => advancePastResolved(importStep)}>
                  Next <ChevronRight className="ml-1 h-4 w-4" />
                </Button>
              )}
            </div>
          </DialogFooter>
        </DialogContent>
      </Dialog>
    </>
  );
}
