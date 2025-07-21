/**
 * Comprehensive Data Processing Tests
 * 
 * This test file validates all implemented functionality for processing different Excel file formats:
 * 1. Traditional mode (separate tax-exempt and taxable columns)
 * 2. Tax type classification mode (single amount column + tax type column)
 * 3. Multi-column sum mode (multiple payment method columns)
 * 4. Merged header processing (complex header structures like Cafe24)
 */

const { ExcelService } = require('../services/ExcelService');
const path = require('path');

describe('Data Processing Integration Tests', () => {
  let excelService;

  beforeEach(() => {
    excelService = new ExcelService();
  });

  describe('Traditional Mode Tests', () => {
    test('should process basic Excel file with separate tax columns', async () => {
      // Test with a basic structure file
      const testOptions = {
        useTaxTypeClassification: false,
        dateColumn: '날짜',
        taxExemptColumn: '면세금액',
        taxableColumn: '과세금액'
      };

      // Mock file path - in real test, you'd use actual test files
      const mockFilePath = 'test-files/basic-structure.xlsx';
      
      // For now, test the logic with mock data
      const mockData = [
        { '날짜': '2024-01-01', '면세금액': 100000, '과세금액': 50000 },
        { '날짜': '2024-01-15', '면세금액': 200000, '과세금액': 75000 },
        { '날짜': '2024-02-01', '면세금액': 150000, '과세금액': 60000 }
      ];

      // Test parseMonthlyData directly
      const result = excelService.parseMonthlyData(mockData, testOptions);
      
      expect(result).toBeDefined();
      expect(result.length).toBeGreaterThan(0);
      
      // Check that data is properly aggregated by month
      const januaryData = result.find(r => r.month === 1 && r.year === 2024);
      expect(januaryData).toBeDefined();
      expect(januaryData.taxExemptAmount).toBe(300000); // 100000 + 200000
      expect(januaryData.taxableAmount).toBe(125000);   // 50000 + 75000
    });

    test('should handle date parsing for various formats', async () => {
      const testDates = [
        '2024.01',
        '2024년 1월',
        '2024-01-15',
        '2024.01.15 10:30:00',
        '2024.01.01 ~ 2024.01.31'
      ];

      testDates.forEach(dateStr => {
        const parsedDate = excelService.parseDate(dateStr);
        expect(parsedDate).toBeDefined();
        expect(parsedDate).toBeInstanceOf(Date);
        expect(parsedDate.getFullYear()).toBe(2024);
      });
    });
  });

  describe('Tax Type Classification Mode Tests', () => {
    test('should process Coupang-style files with tax type classification', async () => {
      const testOptions = {
        useTaxTypeClassification: true,
        useMultiColumnSum: false,
        dateColumn: '매출인식일',
        taxTypeColumn: '과세유형',
        amountColumn: '신용카드(판매)',
        taxExemptValues: ['FREE', '면세', '면세상품'],
        taxableValues: ['TAX', '과세', '과세상품']
      };

      const mockData = [
        { '매출인식일': '2024-12-03', '과세유형': 'FREE', '신용카드(판매)': 24500 },
        { '매출인식일': '2024-12-03', '과세유형': 'TAX', '신용카드(판매)': 15000 },
        { '매출인식일': '2024-11-15', '과세유형': 'FREE', '신용카드(판매)': 30000 },
        { '매출인식일': '2024-11-20', '과세유형': 'TAX', '신용카드(판매)': 20000 }
      ];

      const result = excelService.parseMonthlyDataWithTaxType(mockData, testOptions);
      
      expect(result).toBeDefined();
      expect(result.length).toBeGreaterThan(0);

      // Check December data
      const decemberData = result.find(r => r.month === 12 && r.year === 2024);
      expect(decemberData).toBeDefined();
      expect(decemberData.taxExemptAmount).toBe(24500);
      expect(decemberData.taxableAmount).toBe(15000);

      // Check November data
      const novemberData = result.find(r => r.month === 11 && r.year === 2024);
      expect(novemberData).toBeDefined();
      expect(novemberData.taxExemptAmount).toBe(30000);
      expect(novemberData.taxableAmount).toBe(20000);
    });

    test('should handle unknown tax types by treating as taxable', async () => {
      const testOptions = {
        useTaxTypeClassification: true,
        useMultiColumnSum: false,
        dateColumn: '날짜',
        taxTypeColumn: '과세유형',
        amountColumn: '금액',
        taxExemptValues: ['면세'],
        taxableValues: ['과세']
      };

      const mockData = [
        { '날짜': '2024-01-01', '과세유형': '알수없음', '금액': 10000 }
      ];

      const result = excelService.parseMonthlyDataWithTaxType(mockData, testOptions);
      
      expect(result).toBeDefined();
      expect(result.length).toBe(1);
      expect(result[0].taxExemptAmount).toBe(0);
      expect(result[0].taxableAmount).toBe(10000); // Should be treated as taxable
    });
  });

  describe('Multi-Column Sum Mode Tests', () => {
    test('should sum multiple payment method columns', async () => {
      const testOptions = {
        useTaxTypeClassification: true,
        useMultiColumnSum: true,
        dateColumn: '매출인식일',
        taxTypeColumn: '과세유형',
        amountColumns: ['신용카드(판매)', '현금(판매)', '기타(판매)'],
        taxExemptValues: ['FREE'],
        taxableValues: ['TAX']
      };

      const mockData = [
        { 
          '매출인식일': '2024-12-03', 
          '과세유형': 'FREE', 
          '신용카드(판매)': 10000,
          '현금(판매)': 5000,
          '기타(판매)': 3000
        },
        { 
          '매출인식일': '2024-12-03', 
          '과세유형': 'TAX', 
          '신용카드(판매)': 15000,
          '현금(판매)': 0,
          '기타(판매)': 2000
        }
      ];

      const result = excelService.parseMonthlyDataWithTaxType(mockData, testOptions);
      
      expect(result).toBeDefined();
      expect(result.length).toBe(1);
      
      const decemberData = result[0];
      expect(decemberData.month).toBe(12);
      expect(decemberData.year).toBe(2024);
      expect(decemberData.taxExemptAmount).toBe(18000); // 10000 + 5000 + 3000
      expect(decemberData.taxableAmount).toBe(17000);   // 15000 + 0 + 2000
    });

    test('should handle missing values in multi-column sum', async () => {
      const testOptions = {
        useTaxTypeClassification: true,
        useMultiColumnSum: true,
        dateColumn: '날짜',
        taxTypeColumn: '과세유형',
        amountColumns: ['컬럼1', '컬럼2', '컬럼3'],
        taxExemptValues: ['면세'],
        taxableValues: ['과세']
      };

      const mockData = [
        { 
          '날짜': '2024-01-01', 
          '과세유형': '면세',
          '컬럼1': 10000,
          '컬럼2': null,      // Missing value
          '컬럼3': undefined  // Missing value
        }
      ];

      const result = excelService.parseMonthlyDataWithTaxType(mockData, testOptions);
      
      expect(result).toBeDefined();
      expect(result.length).toBe(1);
      expect(result[0].taxExemptAmount).toBe(10000); // Should handle null/undefined as 0
    });
  });

  describe('Merged Header Processing Tests', () => {
    test('should detect multi-row headers correctly', async () => {
      const mockData = [
        ['거래년월', '서비스구분', '결제수단', '거래건수', '결제금액', ''],
        ['', '', '', '', '과세금액', '면세금액']
      ];

      const result = excelService.detectMultiRowHeaders(mockData, 0);
      
      expect(result.isMultiRow).toBe(true);
      expect(result.headerRows).toEqual([0, 1]);
    });

    test('should combine multi-row headers properly', async () => {
      const mockData = [
        ['거래년월', '서비스구분', '결제수단', '거래건수', '결제금액', ''],
        ['', '', '', '', '과세금액', '면세금액']
      ];

      const result = excelService.combineMultiRowHeaders(mockData, [0, 1]);
      
      expect(result).toContain('거래년월');
      expect(result).toContain('서비스구분');
      expect(result).toContain('결제수단');
      expect(result).toContain('거래건수');
      expect(result).toContain('결제금액 > 과세금액');
      expect(result).toContain('결제금액 > 면세금액');
    });

    test('should process Cafe24-style merged headers', async () => {
      const testOptions = {
        useTaxTypeClassification: false,
        dateColumn: '거래년월',
        taxExemptColumn: '결제금액 > 면세금액',
        taxableColumn: '결제금액 > 과세금액'
      };

      // Mock combined header structure
      const mockData = [
        { 
          '거래년월': '2024-12',
          '서비스구분': '카페24페이먼츠',
          '결제수단': '신용카드',
          '거래건수': 3,
          '결제금액 > 과세금액': 225200,
          '결제금액 > 면세금액': 159500
        },
        { 
          '거래년월': '2024-11',
          '서비스구분': '카페24페이먼츠',
          '결제수단': '신용카드',
          '거래건수': 1,
          '결제금액 > 과세금액': 49200,
          '결제금액 > 면세금액': 0
        }
      ];

      const result = excelService.parseMonthlyData(mockData, testOptions);
      
      expect(result).toBeDefined();
      expect(result.length).toBe(2);

      // Check December data
      const decemberData = result.find(r => r.month === 12);
      expect(decemberData.taxExemptAmount).toBe(159500);
      expect(decemberData.taxableAmount).toBe(225200);

      // Check November data
      const novemberData = result.find(r => r.month === 11);
      expect(novemberData.taxExemptAmount).toBe(0);
      expect(novemberData.taxableAmount).toBe(49200);
    });
  });

  describe('Validation Tests', () => {
    test('should validate required columns for traditional mode', async () => {
      const invalidOptions = {
        useTaxTypeClassification: false,
        dateColumn: '날짜',
        // Missing taxExemptColumn and taxableColumn
      };

      expect(() => {
        excelService.validateProcessingOptions(invalidOptions);
      }).toThrow('면세금액 컬럼과 과세금액 컬럼이 모두 지정되어야 합니다.');
    });

    test('should validate required columns for tax type mode', async () => {
      const invalidOptions = {
        useTaxTypeClassification: true,
        dateColumn: '날짜',
        // Missing taxTypeColumn
      };

      expect(() => {
        excelService.validateProcessingOptions(invalidOptions);
      }).toThrow('과세유형 컬럼이 지정되어야 합니다.');
    });

    test('should validate multi-column sum requirements', async () => {
      const invalidOptions = {
        useTaxTypeClassification: true,
        useMultiColumnSum: true,
        dateColumn: '날짜',
        taxTypeColumn: '과세유형',
        amountColumns: [] // Empty array
      };

      expect(() => {
        excelService.validateProcessingOptions(invalidOptions);
      }).toThrow('다중 컬럼 합계 모드에서는 최소 하나의 금액 컬럼이 지정되어야 합니다.');
    });
  });

  describe('Error Handling Tests', () => {
    test('should handle empty data gracefully', async () => {
      const testOptions = {
        useTaxTypeClassification: false,
        dateColumn: '날짜',
        taxExemptColumn: '면세금액',
        taxableColumn: '과세금액'
      };

      const result = excelService.parseMonthlyData([], testOptions);
      
      expect(result).toBeDefined();
      expect(result).toEqual([]);
    });

    test('should skip rows with invalid dates', async () => {
      const testOptions = {
        useTaxTypeClassification: false,
        dateColumn: '날짜',
        taxExemptColumn: '면세금액',
        taxableColumn: '과세금액'
      };

      const mockData = [
        { '날짜': '2024-01-01', '면세금액': 100000, '과세금액': 50000 },
        { '날짜': '잘못된날짜', '면세금액': 200000, '과세금액': 75000 }, // Invalid date
        { '날짜': null, '면세금액': 150000, '과세금액': 60000 }          // Null date
      ];

      const result = excelService.parseMonthlyData(mockData, testOptions);
      
      expect(result).toBeDefined();
      expect(result.length).toBe(1); // Only valid date row should be processed
      expect(result[0].month).toBe(1);
      expect(result[0].year).toBe(2024);
    });

    test('should handle missing or invalid amount values', async () => {
      const testOptions = {
        useTaxTypeClassification: false,
        dateColumn: '날짜',
        taxExemptColumn: '면세금액',
        taxableColumn: '과세금액'
      };

      const mockData = [
        { '날짜': '2024-01-01', '면세금액': null, '과세금액': '문자열' },
        { '날짜': '2024-01-01', '면세금액': undefined, '과세금액': 50000 }
      ];

      const result = excelService.parseMonthlyData(mockData, testOptions);
      
      expect(result).toBeDefined();
      // Should handle invalid amounts gracefully (convert to 0 or filter out)
    });
  });

  describe('Performance Tests', () => {
    test('should process large datasets efficiently', async () => {
      const testOptions = {
        useTaxTypeClassification: false,
        dateColumn: '날짜',
        taxExemptColumn: '면세금액',
        taxableColumn: '과세금액'
      };

      // Generate large dataset
      const largeDataset = [];
      for (let i = 0; i < 10000; i++) {
        largeDataset.push({
          '날짜': `2024-${String(Math.floor(i / 1000) + 1).padStart(2, '0')}-01`,
          '면세금액': Math.floor(Math.random() * 100000),
          '과세금액': Math.floor(Math.random() * 100000)
        });
      }

      const startTime = Date.now();
      const result = excelService.parseMonthlyData(largeDataset, testOptions);
      const endTime = Date.now();

      expect(result).toBeDefined();
      expect(endTime - startTime).toBeLessThan(1000); // Should complete within 1 second
    });
  });
});

/**
 * Helper function to run actual file tests (requires real Excel files)
 * Uncomment and modify paths to test with real files
 */
/*
describe('Real File Tests', () => {
  const testFilePaths = {
    coupang: '/path/to/coupang/file.xlsx',
    cafe24: '/path/to/cafe24/file.xlsx',
    basic: '/path/to/basic/file.xlsx'
  };

  test('should process real Coupang file', async () => {
    if (!fs.existsSync(testFilePaths.coupang)) {
      return; // Skip if file doesn't exist
    }

    const options = {
      useTaxTypeClassification: true,
      useMultiColumnSum: true,
      dateColumn: '매출인식일',
      taxTypeColumn: '과세유형',
      amountColumns: ['신용카드(판매)', '현금(판매)', '기타(판매)'],
      taxExemptValues: ['FREE'],
      taxableValues: ['TAX']
    };

    const result = await excelService.readExcelFile(testFilePaths.coupang, options);
    
    expect(result).toBeDefined();
    expect(result.length).toBeGreaterThan(0);
  });

  test('should process real Cafe24 file', async () => {
    if (!fs.existsSync(testFilePaths.cafe24)) {
      return; // Skip if file doesn't exist
    }

    const options = {
      useTaxTypeClassification: false,
      dateColumn: '거래년월',
      taxExemptColumn: '결제금액 > 면세금액',
      taxableColumn: '결제금액 > 과세금액'
    };

    const result = await excelService.readExcelFile(testFilePaths.cafe24, options);
    
    expect(result).toBeDefined();
    expect(result.length).toBeGreaterThan(0);
  });
});
*/