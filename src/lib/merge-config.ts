import * as XLSX from 'xlsx';
import type { ExcelFile, WorksheetData, ColumnMapping } from '../components/ExcelCombiner';

// Serialized representation of the merge *decision* — files (by name + order),
// worksheet choices, key columns, column mappings, and toggles. Never carries
// file bytes or row data; data is always re-read when the config is applied.
export interface MergeConfigFile {
  name: string;
}
export interface MergeConfigWorksheet {
  fileIndex: number;
  worksheetName: string;
  headerRow: number;
  keyColumn?: string;
}
export interface MergeConfigMapping {
  outputColumn: string;
  mappings: { fileIndex: number; column: string }[];
}
export interface MergeConfig {
  version: 1;
  files: MergeConfigFile[];
  worksheets: MergeConfigWorksheet[];
  keyColumn?: string;
  columnMappings: MergeConfigMapping[];
  allowIncompleteMappings: boolean;
  allowDoubleMapping: boolean;
}

export interface ConfigPreset {
  id: string;
  name: string;
  config: MergeConfig;
  savedAt: number;
}

const PRESET_KEY = 'merge-config-presets';

export function buildMergeConfig(
  files: ExcelFile[],
  worksheets: WorksheetData[],
  columnMappings: ColumnMapping[],
  keyColumn: string,
  allowIncompleteMappings: boolean,
  allowDoubleMapping: boolean,
): MergeConfig {
  const fileIndex = new Map(files.map((f, i) => [f.id, i]));
  return {
    version: 1,
    files: files.map((f) => ({ name: f.name })),
    worksheets: worksheets.map((w) => ({
      fileIndex: fileIndex.get(w.fileId) ?? 0,
      worksheetName: w.worksheetName,
      headerRow: w.headerRow,
      ...(w.keyColumn ? { keyColumn: w.keyColumn } : {}),
    })),
    keyColumn: keyColumn || undefined,
    columnMappings: columnMappings.map((m) => ({
      outputColumn: m.outputColumn,
      mappings: m.mappings
        .filter((fm) => fileIndex.has(fm.fileId))
        .map((fm) => ({ fileIndex: fileIndex.get(fm.fileId)!, column: fm.column })),
    })),
    allowIncompleteMappings,
    allowDoubleMapping,
  };
}

// Configs are untrusted input (hand-edited files, tampered localStorage) —
// validate shape and coerce types before they become app state.
type UnknownRecord = Record<string, unknown>;

export function sanitizeMergeConfig(raw: unknown): MergeConfig {
  if (!raw || typeof raw !== 'object' || (raw as UnknownRecord).version !== 1) {
    throw new Error('Unsupported config version.');
  }
  const cfg = raw as MergeConfig;
  if (!Array.isArray(cfg.files) || !Array.isArray(cfg.worksheets) || !Array.isArray(cfg.columnMappings)) {
    throw new Error('Invalid config: missing files, worksheets, or mappings.');
  }
  // ponytail: shape+type checks only; full schema validation when configs come from strangers
  return {
    ...cfg,
    files: cfg.files.map((f) => ({ name: String((f as UnknownRecord | null)?.name ?? '') })),
    worksheets: cfg.worksheets.map((w) => {
      const n = Number((w as UnknownRecord | null)?.headerRow);
      return { ...w, headerRow: Number.isFinite(n) && n > 0 ? Math.floor(n) : 1 };
    }),
  };
}

export function parseMergeConfig(json: string): MergeConfig {
  return sanitizeMergeConfig(JSON.parse(json));
}

export function downloadConfig(config: MergeConfig, filename: string) {
  const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json;charset=utf-8;' });
  const url = window.URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  window.URL.revokeObjectURL(url);
}

export function configFilename(config: MergeConfig): string {
  const base = (config.files[0]?.name || 'config').replace(/\.[^.]*$/, '');
  return `merge-config-${base || 'preset'}.json`;
}

export function defaultPresetName(config: MergeConfig): string {
  return (config.files[0]?.name || 'preset').replace(/\.[^.]*$/, '');
}

// --- localStorage presets ---

export function listPresets(): ConfigPreset[] {
  try {
    const raw = localStorage.getItem(PRESET_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function savePreset(name: string, config: MergeConfig): ConfigPreset {
  const presets = listPresets();
  const existing = presets.find((p) => p.name === name);
  const preset: ConfigPreset = {
    id: existing?.id || `${Date.now()}`,
    name,
    config,
    savedAt: Date.now(),
  };
  const next = existing ? presets.map((p) => (p.id === existing.id ? preset : p)) : [preset, ...presets];
  localStorage.setItem(PRESET_KEY, JSON.stringify(next.slice(0, 20)));
  return preset;
}

export function removePreset(id: string) {
  const next = listPresets().filter((p) => p.id !== id);
  localStorage.setItem(PRESET_KEY, JSON.stringify(next));
}

// --- file parsing helpers for the re-point flow ---

export function parseExcelFile(file: File): Promise<string[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        resolve(workbook.SheetNames.filter((s) => s && s.trim() !== ''));
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('Failed to read file'));
    reader.readAsArrayBuffer(file);
  });
}

export async function buildExcelFile(file: File, id: string): Promise<ExcelFile> {
  try {
    const worksheets = await parseExcelFile(file);
    return { id, name: file.name, file, worksheets: worksheets.length ? worksheets : ['Sheet1'] };
  } catch (error) {
    const msg = error instanceof Error ? error.message : String(error);
    const locked = /password|encrypt/i.test(msg);
    return {
      id, name: file.name, file, worksheets: ['Sheet1'],
      readError: locked
        ? 'This file is password-protected. Remove the password in Excel (File → Info → Protect Workbook → Encrypt), re-save, then re-select it here.'
        : 'This file could not be read. Please re-select it from disk.',
    };
  }
}

export function readWorksheetColumns(file: File, worksheetName: string, headerRow = 1): Promise<string[]> {
  return new Promise((resolve) => {
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
        for (let col = range.s.c; col <= range.e.c; col++) {
          const cellAddress = XLSX.utils.encode_cell({ r: headerRow - 1, c: col });
          const cell = worksheet[cellAddress];
          columns.push(cell && cell.v ? String(cell.v) : `Column ${String.fromCharCode(65 + col)}`);
        }
        resolve(columns);
      } catch {
        resolve([]);
      }
    };
    reader.onerror = () => resolve([]);
    reader.readAsArrayBuffer(file);
  });
}
