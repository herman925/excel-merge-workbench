/* eslint-disable */
/**
 * Journey + pentest battery for src/lib/merge-config.ts and src/lib/excel-processor.ts.
 * Pure Node — bundles the real src modules with esbuild; no browser needed.
 *
 *   npm test
 */
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { createRequire } = require('module');
const repoRequire = createRequire(path.join(__dirname, '..', 'package.json'));
const esbuild = repoRequire('esbuild');
const XLSX = repoRequire('xlsx');

const OUT = fs.mkdtempSync(path.join(os.tmpdir(), 'emw-bundle-'));
const pass = [], fail = [];
const t = (name, fn) => { try { fn(); pass.push(name); } catch (e) { fail.push(`${name}: ${e.message}`); } };
const ta = async (name, fn) => { try { await fn(); pass.push(name); } catch (e) { fail.push(`${name}: ${e.message}`); } };
const finish = () => {
  console.log('\n==== RESULTS ====');
  console.log('PASS:', pass.length);
  pass.forEach(p => console.log('  ✓', p));
  console.log('FAIL:', fail.length);
  fail.forEach(f => console.log('  ✗', f));
  process.exit(fail.length ? 1 : 0);
};

(async () => {
  // ---------- Bundle the REAL src modules ----------
  const SRC = path.join(__dirname, '..', 'src', 'lib');
  await esbuild.build({
    entryPoints: [path.join(SRC, 'merge-config.ts'), path.join(SRC, 'excel-processor.ts')],
    bundle: true, format: 'cjs', outdir: OUT, logLevel: 'warning',
  });

  // Node shim: modules use browser localStorage; no-op it in Node.
  const mcPath = path.join(OUT, 'merge-config.js');
  fs.writeFileSync(mcPath,
    'if(typeof localStorage==="undefined"){globalThis.localStorage={getItem:()=>null,setItem:()=>{},removeItem:()=>{}};}\n' +
    fs.readFileSync(mcPath, 'utf8'));

  const mc = require(mcPath);
  const { ExcelProcessor } = require(path.join(OUT, 'excel-processor.js'));

  // Node shim: modules use browser FileReader; back it with fs (test Files carry _path).
  if (typeof FileReader === 'undefined') {
    globalThis.FileReader = class {
      readAsArrayBuffer(file) {
        fs.readFile(file._path, (err, data) => {
          if (err) { if (this.onerror) this.onerror(err); return; }
          this.result = data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength);
          if (this.onload) this.onload({ target: this });
        });
      }
    };
  }
  const mk = (p) => { const f = new File([fs.readFileSync(p)], path.basename(p)); f._path = p; return f; };

  // ---------- Real workbooks ----------
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'emw-'));
  const writeWb = (file, sheets) => {
    const wb = XLSX.utils.book_new();
    for (const [name, rows] of sheets) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);
    XLSX.writeFile(wb, file);
  };

  const f1 = path.join(dir, 'alpha.xlsx');
  writeWb(f1, [['Data', [['ID', 'Name', 'Score'], ['1', 'Alice', 90], ['2', 'Bob', 85], ['3', 'Cara', 70]]]]);
  const f2 = path.join(dir, 'beta.xlsx');
  writeWb(f2, [['Data', [['ID', 'Grade'], ['1', 'A'], ['2', 'B-'], ['4', 'C']]]]);
  const f3 = path.join(dir, 'gamma.xlsx'); // two sheets
  writeWb(f3, [['Cover', [['junk'], ['x']]], ['Data', [['ID', 'Note'], ['1', 'hello']]]]);
  const fBad = path.join(dir, 'bad.xlsx');
  fs.writeFileSync(fBad, Buffer.from('this is not an excel file'));
  const fEnc = path.join(dir, 'enc.xlsx');
  fs.writeFileSync(fEnc, Buffer.concat([Buffer.from('PK\x03\x04'), Buffer.alloc(512, 0x41)]));

  // ---------- PENTEST: config sanitization ----------
  t('sanitize: rejects version 2', () => assert.throws(() => mc.sanitizeMergeConfig({ version: 2, files: [], worksheets: [], columnMappings: [] })));
  t('sanitize: rejects missing arrays', () => assert.throws(() => mc.sanitizeMergeConfig({ version: 1, files: 'x', worksheets: [], columnMappings: [] })));
  t('sanitize: rejects null', () => assert.throws(() => mc.sanitizeMergeConfig(null)));
  t('sanitize: headerRow "3" -> 3', () => assert.strictEqual(mc.sanitizeMergeConfig({ version: 1, files: [], worksheets: [{ headerRow: '3' }], columnMappings: [] }).worksheets[0].headerRow, 3));
  t('sanitize: headerRow -5 -> 1', () => assert.strictEqual(mc.sanitizeMergeConfig({ version: 1, files: [], worksheets: [{ headerRow: -5 }], columnMappings: [] }).worksheets[0].headerRow, 1));
  t('sanitize: headerRow Infinity -> 1', () => assert.strictEqual(mc.sanitizeMergeConfig({ version: 1, files: [], worksheets: [{ headerRow: Infinity }], columnMappings: [] }).worksheets[0].headerRow, 1));
  t('sanitize: file name 42 -> "42"', () => assert.strictEqual(mc.sanitizeMergeConfig({ version: 1, files: [{ name: 42 }], worksheets: [], columnMappings: [] }).files[0].name, '42'));
  t('parse: garbage JSON throws', () => assert.throws(() => mc.parseMergeConfig('{oops')));
  t('parse: deep-nested bomb throws', () => {
    let s = '[]';
    for (let i = 0; i < 100000; i++) s = '[' + s + ']';
    assert.throws(() => mc.parseMergeConfig(s));
  });
  t('parse: legit round-trip keeps keys', () => {
    const cfg = mc.buildMergeConfig(
      [{ id: 'f1', name: 'alpha.xlsx' }, { id: 'f2', name: 'beta.xlsx' }],
      [
        { fileId: 'f1', worksheetName: 'Data', headerRow: 1, columns: ['ID'], keyColumn: 'ID' },
        { fileId: 'f2', worksheetName: 'Data', headerRow: 1, columns: ['ID'], keyColumn: 'ID' },
      ],
      [{ outputColumn: 'ID', mappings: [{ fileId: 'f1', column: 'ID' }, { fileId: 'f2', column: 'ID' }] }],
      'ID', false, false
    );
    const back = mc.parseMergeConfig(JSON.stringify(cfg));
    assert.strictEqual(back.version, 1);
    assert.strictEqual(back.worksheets[0].keyColumn, 'ID');
    assert.strictEqual(back.columnMappings[0].mappings.length, 2);
  });

  // ---------- USER JOURNEY ----------
  await ta('journey: alpha parses, sheets=["Data"]', async () => {
    const ef = await mc.buildExcelFile(mk(f1), 'x1');
    assert.ok(!ef.readError);
    assert.deepStrictEqual(ef.worksheets, ['Data']);
  });

  await ta('journey: gamma lists both sheets', async () => {
    const ef = await mc.buildExcelFile(mk(f3), 'x3');
    assert.ok(!ef.readError);
    assert.deepStrictEqual(ef.worksheets, ['Cover', 'Data']);
  });

  await ta('journey: columns re-read from picked sheet', async () => {
    const cols = await mc.readWorksheetColumns(mk(f3), 'Data', 1);
    assert.deepStrictEqual(cols, ['ID', 'Note']);
  });

  await ta('pentest: malformed file does not crash (SheetJS sniffs text)', async () => {
    const ef = await mc.buildExcelFile(mk(fBad), 'xb');
    assert.ok(Array.isArray(ef.worksheets) && ef.worksheets.length > 0);
  });

  await ta('pentest: garbage bytes with zip magic do not crash', async () => {
    const ef = await mc.buildExcelFile(mk(fEnc), 'xe');
    assert.ok(typeof ef.readError === 'string' || !ef.readError);
  });

  await ta('journey: FULL MERGE via real processor', async () => {
    const efA = await mc.buildExcelFile(mk(f1), 'x1');
    const efB = await mc.buildExcelFile(mk(f2), 'x2');
    const files = [efA, efB];
    const worksheets = [
      { fileId: 'x1', worksheetName: 'Data', headerRow: 1, columns: ['ID', 'Name', 'Score'], keyColumn: 'ID' },
      { fileId: 'x2', worksheetName: 'Data', headerRow: 1, columns: ['ID', 'Grade'], keyColumn: 'ID' },
    ];
    const mappings = [{ outputColumn: 'ID', mappings: [{ fileId: 'x1', column: 'ID' }, { fileId: 'x2', column: 'ID' }] }];
    const proc = new ExcelProcessor(files, worksheets, mappings, '');
    const res = await proc.processFiles();
    assert.strictEqual(res.totalRowsProcessed, 6, '3+3 rows read');
    assert.strictEqual(res.combinedData.length - 1, 4, 'keys 1,2,3,4 -> 4 rows');
    const blob = ExcelProcessor.generateCSVWithBOM(res.combinedData);
    const text = await blob.text();
    const lines = text.split('\n');
    assert.strictEqual(lines.length, 5, 'header + 4 rows');
    assert.ok(lines.includes('1'), 'ID 1 present in CSV');

    const cfg = mc.buildMergeConfig(files, worksheets, mappings, 'ID', false, false);
    const back = mc.parseMergeConfig(JSON.stringify(cfg));
    assert.strictEqual(back.files.length, 2);
  });

  await ta('pentest: corrupted config mid-journey recovers via sanitize', async () => {
    const cfg = mc.buildMergeConfig(
      [{ id: 'f1', name: 'alpha.xlsx' }],
      [{ fileId: 'f1', worksheetName: 'Data', headerRow: 2, columns: ['ID'] }],
      [], '', false, false
    );
    cfg.worksheets[0].headerRow = '9'; // simulate tampering
    const back = mc.sanitizeMergeConfig(cfg);
    assert.strictEqual(back.worksheets[0].headerRow, 9);
  });

  finish();
})().catch(e => { fail.push('driver: ' + e.message); finish(); });
