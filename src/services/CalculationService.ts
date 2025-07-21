import { ShoppingMallData, CalculationResult, MonthlyTotal } from '../types';

export class CalculationService {
  calculateTotals(mallData: ShoppingMallData): CalculationResult {
    // Group data by year and month
    const groupedData = this.groupByYearMonth(mallData.data);
    
    // Calculate monthly totals
    const monthlyTotals: MonthlyTotal[] = [];
    let yearlyTaxExempt = 0;
    let yearlyTaxable = 0;

    Object.entries(groupedData).forEach(([yearMonth, items]) => {
      const [year, month] = yearMonth.split('-').map(Number);
      
      const monthlyTaxExempt = items.reduce((sum, item) => sum + item.taxExemptAmount, 0);
      const monthlyTaxable = items.reduce((sum, item) => sum + item.taxableAmount, 0);
      
      monthlyTotals.push({
        year,
        month,
        taxExempt: monthlyTaxExempt,
        taxable: monthlyTaxable,
        total: monthlyTaxExempt + monthlyTaxable
      });

      yearlyTaxExempt += monthlyTaxExempt;
      yearlyTaxable += monthlyTaxable;
    });

    // Sort by year and month
    monthlyTotals.sort((a, b) => {
      if (a.year !== b.year) return a.year - b.year;
      return a.month - b.month;
    });

    return {
      mallName: mallData.mallName,
      monthlyTotals,
      yearlyTotal: {
        taxExempt: yearlyTaxExempt,
        taxable: yearlyTaxable,
        total: yearlyTaxExempt + yearlyTaxable
      }
    };
  }

  private groupByYearMonth(data: any[]): Record<string, any[]> {
    return data.reduce((groups, item) => {
      const key = `${item.year}-${item.month}`;
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(item);
      return groups;
    }, {} as Record<string, any[]>);
  }

  formatCurrency(amount: number): string {
    return new Intl.NumberFormat('ko-KR', {
      style: 'currency',
      currency: 'KRW'
    }).format(amount);
  }

  calculateGrowthRate(current: number, previous: number): number {
    if (previous === 0) return 0;
    return ((current - previous) / previous) * 100;
  }
}