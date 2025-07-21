import { ExcelService } from '../services/ExcelService';
import * as fs from 'fs';
import * as path from 'path';

// Mock xlsx module
jest.mock('xlsx', () => ({
  readFile: jest.fn(),
  utils: {
    sheet_to_json: jest.fn(),
    json_to_sheet: jest.fn(),
    book_new: jest.fn(() => ({})),
    book_append_sheet: jest.fn()
  },
  writeFile: jest.fn()
}));

jest.mock('fs');

describe('ExcelService', () => {
  let excelService: ExcelService;
  const XLSX = require('xlsx');

  beforeEach(() => {
    excelService = new ExcelService();
    jest.clearAllMocks();
  });

  describe('parseDate', () => {
    it('should parse Excel serial date', () => {
      // Excel serial date for 2024-01-15
      const serialDate = 45306;
      const date = (excelService as any).parseDate(serialDate);
      
      expect(date.getFullYear()).toBe(2024);
      expect(date.getMonth()).toBe(0); // January is 0
      expect(date.getDate()).toBe(15);
    });

    it('should parse Korean date format', () => {
      const koreanDate = '2024년 3월 15일';
      const date = (excelService as any).parseDate(koreanDate);
      
      expect(date.getFullYear()).toBe(2024);
      expect(date.getMonth()).toBe(2); // March is 2
      expect(date.getDate()).toBe(15);
    });

    it('should parse standard date string', () => {
      const dateString = '2024-03-15';
      const date = (excelService as any).parseDate(dateString);
      
      expect(date.getFullYear()).toBe(2024);
      expect(date.getMonth()).toBe(2); // March is 2
      expect(date.getDate()).toBe(15);
    });

    it('should handle Date object', () => {
      const dateObj = new Date(2024, 2, 15); // March 15, 2024
      const date = (excelService as any).parseDate(dateObj);
      
      expect(date).toBe(dateObj);
    });

    it('should throw error for invalid date', () => {
      expect(() => {
        (excelService as any).parseDate('invalid date');
      }).toThrow('Unable to parse date: invalid date');
    });
  });

  describe('parseAmount', () => {
    it('should parse numeric value', () => {
      const amount = (excelService as any).parseAmount(1234567);
      expect(amount).toBe(1234567);
    });

    it('should parse string with currency symbol and commas', () => {
      const amount = (excelService as any).parseAmount('₩1,234,567');
      expect(amount).toBe(1234567);
    });

    it('should parse string with dollar sign', () => {
      const amount = (excelService as any).parseAmount('$1,234,567');
      expect(amount).toBe(1234567);
    });

    it('should handle empty string', () => {
      const amount = (excelService as any).parseAmount('');
      expect(amount).toBe(0);
    });

    it('should handle null/undefined', () => {
      expect((excelService as any).parseAmount(null)).toBe(0);
      expect((excelService as any).parseAmount(undefined)).toBe(0);
    });
  });

  describe('analyzeExcelFile', () => {
    it('should provide comprehensive analysis of Excel file', async () => {
      const mockWorkbook = {
        SheetNames: ['Sheet1'],
        Sheets: {
          Sheet1: {
            '!merges': []
          }
        }
      };

      const mockData = [
        ['Report Title'],
        [],
        ['날짜', '면세금액', '과세금액'],
        ['2024-01-01', 100000, 50000],
        ['2024-01-02', 200000, 75000]
      ];

      // Mock file stats
      (fs.statSync as jest.Mock).mockReturnValue({ size: 1024 });
      
      // Mock XLSX methods
      (XLSX.readFile as jest.Mock).mockReturnValue(mockWorkbook);
      (XLSX.utils.sheet_to_json as jest.Mock).mockReturnValue(mockData);

      const analysis = await excelService.analyzeExcelFile('test.xlsx');

      expect(analysis.fileInfo.name).toBe('test.xlsx');
      expect(analysis.fileInfo.size).toBe(1024);
      expect(analysis.headerDetection.detectedRow).toBe(2);
      expect(analysis.headerDetection.confidence).toBeGreaterThan(80);
      expect(analysis.headerDetection.headers).toEqual(['날짜', '면세금액', '과세금액']);
      expect(analysis.structure.totalRows).toBe(5);
      expect(analysis.dataPreview.firstDataRow).toBe(3);
      expect(analysis.dataPreview.isEmpty).toBe(false);
      expect(analysis.recommendations.length).toBeGreaterThan(0);
    });
  });

  describe('readExcelFile', () => {
    it('should read and parse Excel file data', async () => {
      const mockWorkbook = {
        SheetNames: ['Sheet1', 'Sheet2'],
        Sheets: {
          Sheet1: {}
        }
      };

      const mockData = [
        ['날짜', '면세금액', '과세금액'],
        ['2024년 1월 15일', '₩1,000,000', '₩500,000'],
        ['2024년 2월 20일', '₩2,000,000', '₩1,000,000']
      ];

      XLSX.readFile.mockReturnValue(mockWorkbook);
      XLSX.utils.sheet_to_json.mockReturnValue(mockData);

      const result = await excelService.readExcelFile('/test/file.xlsx', {});

      expect(result).toHaveLength(2);
      expect(result[0]).toEqual({
        month: 1,
        year: 2024,
        taxExemptAmount: 1000000,
        taxableAmount: 500000
      });
      expect(result[1]).toEqual({
        month: 2,
        year: 2024,
        taxExemptAmount: 2000000,
        taxableAmount: 1000000
      });
    });

    it('should use custom column names', async () => {
      const mockWorkbook = {
        SheetNames: ['Sheet1'],
        Sheets: {
          Sheet1: {}
        }
      };

      const mockData = [
        ['Date', 'Tax Free', 'Taxable'],
        ['2024-01-15', 1000000, 500000]
      ];

      XLSX.readFile.mockReturnValue(mockWorkbook);
      XLSX.utils.sheet_to_json.mockReturnValue(mockData);

      const result = await excelService.readExcelFile('/test/file.xlsx', {
        dateColumn: 'Date',
        taxExemptColumn: 'Tax Free',
        taxableColumn: 'Taxable'
      });

      expect(result).toHaveLength(1);
      expect(result[0]).toEqual({
        month: 1,
        year: 2024,
        taxExemptAmount: 1000000,
        taxableAmount: 500000
      });
    });

    it('should throw error if sheet not found', async () => {
      const mockWorkbook = {
        SheetNames: ['Sheet1'],
        Sheets: {
          Sheet1: {}
        }
      };

      XLSX.readFile.mockReturnValue(mockWorkbook);

      await expect(
        excelService.readExcelFile('/test/file.xlsx', { sheetName: 'NonExistentSheet' })
      ).rejects.toThrow('Sheet NonExistentSheet not found in file');
    });
  });
});