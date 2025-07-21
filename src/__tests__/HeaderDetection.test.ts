import { ExcelService } from '../services/ExcelService';
import * as XLSX from 'xlsx';
import * as path from 'path';
import * as fs from 'fs';

describe('ExcelService Header Detection', () => {
  let excelService: ExcelService;
  const testDir = path.join(__dirname, '..', '..', 'test-data');

  beforeAll(() => {
    excelService = new ExcelService();
    
    // Create test directory
    if (!fs.existsSync(testDir)) {
      fs.mkdirSync(testDir, { recursive: true });
    }
  });

  afterAll(() => {
    // Clean up test files
    if (fs.existsSync(testDir)) {
      fs.readdirSync(testDir).forEach(file => {
        fs.unlinkSync(path.join(testDir, file));
      });
      fs.rmdirSync(testDir);
    }
  });

  const createTestFile = (filename: string, data: any[][], merges?: any[]) => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    if (merges) {
      ws['!merges'] = merges;
    }
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    const filePath = path.join(testDir, filename);
    XLSX.writeFile(wb, filePath);
    return filePath;
  };

  describe('detectHeaderRow', () => {
    it('should detect headers in first row', async () => {
      const filePath = createTestFile('headers-first-row.xlsx', [
        ['Date', 'Tax Exempt', 'Taxable', 'Total'],
        ['2024-01-01', 100000, 50000, 150000],
        ['2024-01-02', 200000, 75000, 275000],
      ]);

      const result = await excelService.detectHeaderRow(filePath);
      
      expect(result.headerRow).toBe(0);
      expect(result.confidence).toBeGreaterThanOrEqual(80);
      expect(result.reasons).toContain('High text ratio: 100%');
    });

    it('should detect headers after empty rows', async () => {
      const filePath = createTestFile('headers-after-empty.xlsx', [
        ['Shopping Mall Sales Report'],
        [],
        ['날짜', '면세금액', '과세금액', '총액'],
        ['2024-01-01', 100000, 50000, 150000],
        ['2024-01-02', 200000, 75000, 275000],
      ]);

      const result = await excelService.detectHeaderRow(filePath);
      
      expect(result.headerRow).toBe(2);
      expect(result.confidence).toBeGreaterThanOrEqual(80);
      expect(result.reasons).toContain('Preceded by empty row');
    });

    it('should detect Korean headers', async () => {
      const filePath = createTestFile('korean-headers.xlsx', [
        ['쇼핑몰 매출 데이터'],
        ['기간: 2024년 1월'],
        [],
        ['날짜', '주문번호', '면세금액', '과세금액', '배송비', '총액'],
        ['2024-01-01', 'ORD001', 100000, 50000, 2500, 152500],
      ]);

      const result = await excelService.detectHeaderRow(filePath);
      
      expect(result.headerRow).toBe(3);
      expect(result.confidence).toBeGreaterThanOrEqual(80);
      expect(result.reasons.some(r => r.includes('header keywords'))).toBe(true);
    });

    it('should handle files with no clear headers', async () => {
      const filePath = createTestFile('no-headers.xlsx', [
        [100, 200, 300],
        [150, 250, 350],
        [200, 300, 400],
      ]);

      const result = await excelService.detectHeaderRow(filePath);
      
      expect(result.headerRow).toBe(0);
      // Files with all numeric data still get some score for having non-empty cells
      expect(result.confidence).toBeLessThanOrEqual(50);
      // Check that no header keywords were found
      expect(result.reasons.some(r => r.includes('header keywords'))).toBe(false);
    });

    it('should detect headers by pattern change', async () => {
      const filePath = createTestFile('pattern-change.xlsx', [
        ['Column A', 'Column B', 'Column C'],
        [123, 456, 789],
        [111, 222, 333],
      ]);

      const result = await excelService.detectHeaderRow(filePath);
      
      expect(result.headerRow).toBe(0);
      expect(result.reasons.some(r => r.includes('Pattern change'))).toBe(true);
    });
  });

  describe('getColumnHeaders with auto-detection', () => {
    it('should auto-detect and retrieve headers', async () => {
      const filePath = createTestFile('auto-detect-headers.xlsx', [
        ['Report Title'],
        [],
        ['Date', 'Amount', 'Type'],
        ['2024-01-01', 1000, 'Sale'],
      ]);

      const headers = await excelService.getColumnHeaders(filePath);
      
      expect(headers).toEqual(['Date', 'Amount', 'Type']);
    });

    it('should handle merged cells in headers', async () => {
      const filePath = createTestFile('merged-headers.xlsx', [
        ['Sales Data', '', 'Amounts', ''],
        ['Date', 'Type', 'Tax Exempt', 'Taxable'],
        ['2024-01-01', 'Online', 100000, 50000],
      ], [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } },
        { s: { r: 0, c: 2 }, e: { r: 0, c: 3 } },
      ]);

      const headers = await excelService.getColumnHeaders(filePath, { headerRow: 1 });
      
      expect(headers).toEqual(['Date', 'Type', 'Tax Exempt', 'Taxable']);
    });

    it('should allow manual header row specification', async () => {
      const filePath = createTestFile('manual-header-row.xlsx', [
        ['Title'],
        ['Subtitle'],
        ['Col1', 'Col2', 'Col3'],
        ['Data1', 'Data2', 'Data3'],
      ]);

      const headers = await excelService.getColumnHeaders(filePath, { 
        headerRow: 2,
        autoDetect: false 
      });
      
      expect(headers).toEqual(['Col1', 'Col2', 'Col3']);
    });
  });

  describe('readExcelFile with dynamic headers', () => {
    it('should read data with auto-detected headers', async () => {
      const filePath = createTestFile('read-with-auto-headers.xlsx', [
        ['쇼핑몰 A'],
        [],
        ['날짜', '면세금액', '과세금액'],
        ['2024-01-01', '100,000', '50,000'],
        ['2024-01-02', '200,000', '75,000'],
      ]);

      const data = await excelService.readExcelFile(filePath, {
        dateColumn: '날짜',
        taxExemptColumn: '면세금액',
        taxableColumn: '과세금액'
      });

      expect(data).toHaveLength(2);
      expect(data[0]).toEqual({
        month: 1,
        year: 2024,
        taxExemptAmount: 100000,
        taxableAmount: 50000
      });
    });

    it('should handle complex multi-row data', async () => {
      const filePath = createTestFile('complex-data.xlsx', [
        ['Shopping Mall Report 2024'],
        ['Generated: 2024-01-15'],
        [],
        [],
        ['Date', 'Order ID', 'Product', 'Tax Exempt Amount', 'Taxable Amount', 'Shipping', 'Total'],
        ['2024-01-01', 'ORD001', 'Product A', '100,000', '50,000', '2,500', '152,500'],
        ['2024-01-02', 'ORD002', 'Product B', '200,000', '75,000', '0', '275,000'],
      ]);

      const data = await excelService.readExcelFile(filePath, {
        dateColumn: 'Date',
        taxExemptColumn: 'Tax Exempt Amount',
        taxableColumn: 'Taxable Amount'
      });

      expect(data).toHaveLength(2);
      expect(data[0].taxExemptAmount).toBe(100000);
      expect(data[1].taxableAmount).toBe(75000);
    });
  });
});