import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { ExcelService } from '../services/ExcelService';
import { CalculationService } from '../services/CalculationService';

describe('official SheetJS real-file regression', () => {
  let dir: string;
  beforeEach(() => { dir = fs.mkdtempSync(path.join(os.tmpdir(), 'tax-sheetjs-')); });
  afterEach(() => fs.rmSync(dir, { recursive: true, force: true }));

  it.each(['xls', 'xlsx'])('imports %s, detects headers, calculates and exports', async extension => {
    expect(XLSX.version).toBe('0.20.3');
    const input = path.join(dir, `input.${extension}`);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([
      ['날짜', '면세금액', '과세금액'],
      ['2026-01-10', 1000, 2000],
      ['2026-01-20', 3000, 4000],
    ]), 'Sales');
    XLSX.writeFile(workbook, input);
    const service = new ExcelService();
    expect(await service.getColumnHeaders(input)).toEqual(expect.arrayContaining(['날짜', '면세금액', '과세금액']));
    const data = await service.readExcelFile(input, { dateColumn: '날짜', taxExemptColumn: '면세금액', taxableColumn: '과세금액' });
    const result = new CalculationService().calculateTotals({ mallName: 'Synthetic', filePath: input, data });
    expect(result.yearlyTotal).toEqual({ taxExempt: 4000, taxable: 6000, total: 10000 });
    const output = path.join(dir, 'output.xlsx');
    await service.exportResults([result], output, 'xlsx');
    const exported = XLSX.readFile(output);
    expect(exported.SheetNames).toEqual(['요약', 'Synthetic']);
    expect(XLSX.utils.sheet_to_json(exported.Sheets['요약'])).toEqual([{
      '쇼핑몰': 'Synthetic', '연간 면세 합계': 4000, '연간 과세 합계': 6000, '연간 총 합계': 10000,
    }]);
    const csv = path.join(dir, 'output.csv');
    await service.exportResults([result], csv, 'csv');
    expect(fs.readFileSync(csv, 'utf8')).toContain('Synthetic,4000,6000,10000');
  });
});
