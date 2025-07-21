export interface ShoppingMallData {
  mallName: string;
  filePath: string;
  data: MonthlyData[];
}

export interface MonthlyData {
  month: number;
  year: number;
  taxExemptAmount: number;
  taxableAmount: number;
}

export interface CalculationResult {
  mallName: string;
  monthlyTotals: MonthlyTotal[];
  yearlyTotal: {
    taxExempt: number;
    taxable: number;
    total: number;
  };
}

export interface MonthlyTotal {
  month: number;
  year: number;
  taxExempt: number;
  taxable: number;
  total: number;
}

export interface ProcessingOptions {
  dateColumn?: string;
  taxExemptColumn?: string;
  taxableColumn?: string;
  sheetName?: string;
  headerRow?: number;  // Optional: specify header row index, otherwise auto-detect
  
  // New options for tax type based processing
  useTaxTypeClassification?: boolean;  // Whether to use tax type classification
  taxTypeColumn?: string;              // Column that contains tax type (과세유형)
  amountColumn?: string;               // Single amount column when using tax type classification
  taxExemptValues?: string[];          // Values that indicate tax-exempt (면세)
  taxableValues?: string[];            // Values that indicate taxable (과세)
  
  // Multi-column sum options for complex payment structures (like Coupang)
  useMultiColumnSum?: boolean;         // Whether to sum multiple amount columns
  amountColumns?: string[];            // Array of amount columns to sum together
}