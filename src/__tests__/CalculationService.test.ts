import { CalculationService } from '../services/CalculationService';
import { ShoppingMallData, MonthlyData } from '../types';

describe('CalculationService', () => {
  let calculationService: CalculationService;

  beforeEach(() => {
    calculationService = new CalculationService();
  });

  describe('calculateTotals', () => {
    it('should calculate monthly and yearly totals correctly', () => {
      const mockData: MonthlyData[] = [
        { month: 1, year: 2024, taxExemptAmount: 1000000, taxableAmount: 500000 },
        { month: 1, year: 2024, taxExemptAmount: 2000000, taxableAmount: 1000000 },
        { month: 2, year: 2024, taxExemptAmount: 1500000, taxableAmount: 750000 },
        { month: 3, year: 2024, taxExemptAmount: 1800000, taxableAmount: 900000 },
      ];

      const mallData: ShoppingMallData = {
        mallName: 'Test Mall',
        filePath: '/test/path.xlsx',
        data: mockData
      };

      const result = calculationService.calculateTotals(mallData);

      expect(result.mallName).toBe('Test Mall');
      expect(result.monthlyTotals).toHaveLength(3);
      
      // Check January totals (sum of two entries)
      const januaryTotal = result.monthlyTotals.find(m => m.month === 1 && m.year === 2024);
      expect(januaryTotal).toBeDefined();
      expect(januaryTotal!.taxExempt).toBe(3000000);
      expect(januaryTotal!.taxable).toBe(1500000);
      expect(januaryTotal!.total).toBe(4500000);

      // Check yearly totals
      expect(result.yearlyTotal.taxExempt).toBe(6300000);
      expect(result.yearlyTotal.taxable).toBe(3150000);
      expect(result.yearlyTotal.total).toBe(9450000);
    });

    it('should handle empty data', () => {
      const mallData: ShoppingMallData = {
        mallName: 'Empty Mall',
        filePath: '/test/empty.xlsx',
        data: []
      };

      const result = calculationService.calculateTotals(mallData);

      expect(result.mallName).toBe('Empty Mall');
      expect(result.monthlyTotals).toHaveLength(0);
      expect(result.yearlyTotal.taxExempt).toBe(0);
      expect(result.yearlyTotal.taxable).toBe(0);
      expect(result.yearlyTotal.total).toBe(0);
    });

    it('should sort monthly totals by year and month', () => {
      const mockData: MonthlyData[] = [
        { month: 3, year: 2024, taxExemptAmount: 1000000, taxableAmount: 500000 },
        { month: 1, year: 2024, taxExemptAmount: 2000000, taxableAmount: 1000000 },
        { month: 2, year: 2023, taxExemptAmount: 1500000, taxableAmount: 750000 },
        { month: 12, year: 2023, taxExemptAmount: 1800000, taxableAmount: 900000 },
      ];

      const mallData: ShoppingMallData = {
        mallName: 'Test Mall',
        filePath: '/test/path.xlsx',
        data: mockData
      };

      const result = calculationService.calculateTotals(mallData);

      expect(result.monthlyTotals[0]).toMatchObject({ year: 2023, month: 2 });
      expect(result.monthlyTotals[1]).toMatchObject({ year: 2023, month: 12 });
      expect(result.monthlyTotals[2]).toMatchObject({ year: 2024, month: 1 });
      expect(result.monthlyTotals[3]).toMatchObject({ year: 2024, month: 3 });
    });
  });

  describe('formatCurrency', () => {
    it('should format currency in Korean Won', () => {
      const formatted = calculationService.formatCurrency(1234567);
      expect(formatted).toContain('1,234,567');
      expect(formatted).toMatch(/₩|KRW/);
    });

    it('should handle zero amount', () => {
      const formatted = calculationService.formatCurrency(0);
      expect(formatted).toMatch(/₩0|KRW\s*0/);
    });
  });

  describe('calculateGrowthRate', () => {
    it('should calculate positive growth rate', () => {
      const rate = calculationService.calculateGrowthRate(150, 100);
      expect(rate).toBe(50);
    });

    it('should calculate negative growth rate', () => {
      const rate = calculationService.calculateGrowthRate(50, 100);
      expect(rate).toBe(-50);
    });

    it('should handle zero previous value', () => {
      const rate = calculationService.calculateGrowthRate(100, 0);
      expect(rate).toBe(0);
    });

    it('should handle same values', () => {
      const rate = calculationService.calculateGrowthRate(100, 100);
      expect(rate).toBe(0);
    });
  });
});