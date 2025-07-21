import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';
import { MonthlyData, ProcessingOptions } from '../types';

export class ExcelService {
  async readExcelFile(filePath: string, options: ProcessingOptions): Promise<MonthlyData[]> {
    const workbook = XLSX.readFile(filePath);
    const sheetName = options.sheetName || workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    if (!worksheet) {
      throw new Error(`Sheet ${sheetName} not found in file`);
    }

    // Detect header row automatically if not specified
    let headerRowIndex = 0;
    let isMultiRowHeader = false;
    let headerRows: number[] = [];
    
    if (options.headerRow === undefined) {
      const detection = await this.detectHeaderRow(filePath, sheetName);
      headerRowIndex = detection.headerRow;
      isMultiRowHeader = detection.isMultiRowHeader || false;
      headerRows = detection.headerRows || [headerRowIndex];
      console.log(`Auto-detected header at row ${headerRowIndex} with ${detection.confidence}% confidence`);
      if (isMultiRowHeader) {
        console.log(`Multi-row header detected: rows ${headerRows.join(', ')}`);
      }
    } else {
      headerRowIndex = options.headerRow;
      headerRows = [headerRowIndex];
    }

    // Check if we need to handle merged cells
    const merges = worksheet['!merges'];
    let data: any[];
    
    if (merges && merges.length > 0) {
      // Use array of arrays for better merged cell handling
      const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
      
      // Process merged cells - fill in the blanks
      const processedData = this.processMergedCells(rawData as any[][], merges);
      
      // Convert back to object format using detected header row(s)
      if (processedData.length > Math.max(...headerRows)) {
        let headers: any[];
        let dataStartRow: number;
        
        if (isMultiRowHeader && headerRows.length > 1) {
          // Combine multi-row headers
          const combinedHeaders = this.combineMultiRowHeaders(processedData, headerRows);
          headers = combinedHeaders;
          dataStartRow = Math.max(...headerRows) + 1;
        } else {
          headers = processedData[headerRowIndex];
          dataStartRow = headerRowIndex + 1;
        }
        
        data = processedData.slice(dataStartRow).map(row => {
          const obj: any = {};
          headers.forEach((header: any, index: number) => {
            if (header) {
              obj[header] = row[index];
            }
          });
          return obj;
        });
      } else {
        data = [];
      }
    } else {
      // No merged cells, use standard approach with dynamic header row
      const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false });
      
      if (rawData.length > Math.max(...headerRows)) {
        let headers: any[];
        let dataStartRow: number;
        
        if (isMultiRowHeader && headerRows.length > 1) {
          // Combine multi-row headers
          const combinedHeaders = this.combineMultiRowHeaders(rawData as any[][], headerRows);
          headers = combinedHeaders;
          dataStartRow = Math.max(...headerRows) + 1;
        } else {
          headers = rawData[headerRowIndex] as any[];
          dataStartRow = headerRowIndex + 1;
        }
        
        data = rawData.slice(dataStartRow).map(row => {
          const obj: any = {};
          headers.forEach((header: any, index: number) => {
            if (header) {
              obj[header] = (row as any[])[index];
            }
          });
          return obj;
        });
      } else {
        data = [];
      }
    }
    
    // Validate that we have data to process
    if (!data || data.length === 0) {
      console.warn('No data found after header row', headerRowIndex);
      return [];
    }
    
    // Validate column names exist based on processing mode
    if (!options.dateColumn) {
      throw new Error('날짜 컬럼이 지정되지 않았습니다.');
    }
    
    if (options.useTaxTypeClassification) {
      // Tax type classification mode validation
      if (!options.taxTypeColumn) {
        throw new Error('과세유형 컬럼이 지정되어야 합니다.');
      }
      
      if (options.useMultiColumnSum) {
        if (!options.amountColumns || options.amountColumns.length === 0) {
          throw new Error('다중 컬럼 합계 모드에서는 최소 하나의 금액 컬럼이 지정되어야 합니다.');
        }
      } else {
        if (!options.amountColumn) {
          throw new Error('단일 컬럼 모드에서는 금액 컬럼이 지정되어야 합니다.');
        }
      }
    } else {
      // Traditional mode validation
      if (!options.taxExemptColumn || !options.taxableColumn) {
        throw new Error('면세금액 컬럼과 과세금액 컬럼이 모두 지정되어야 합니다.');
      }
    }
    
    // Check if the specified columns exist in the data
    const sampleRow = data[0];
    if (sampleRow && typeof sampleRow === 'object') {
      const availableColumns = Object.keys(sampleRow);
      const missingColumns = [];
      
      if (!availableColumns.includes(options.dateColumn)) {
        missingColumns.push(`날짜 컬럼 '${options.dateColumn}'`);
      }
      
      if (options.useTaxTypeClassification) {
        if (!availableColumns.includes(options.taxTypeColumn!)) {
          missingColumns.push(`과세유형 컬럼 '${options.taxTypeColumn}'`);
        }
        
        if (options.useMultiColumnSum) {
          // Check all amount columns exist
          options.amountColumns?.forEach(colName => {
            if (!availableColumns.includes(colName)) {
              missingColumns.push(`금액 컬럼 '${colName}'`);
            }
          });
        } else {
          // Check single amount column
          if (!availableColumns.includes(options.amountColumn!)) {
            missingColumns.push(`금액 컬럼 '${options.amountColumn}'`);
          }
        }
      } else {
        if (!availableColumns.includes(options.taxExemptColumn!)) {
          missingColumns.push(`면세금액 컬럼 '${options.taxExemptColumn}'`);
        }
        if (!availableColumns.includes(options.taxableColumn!)) {
          missingColumns.push(`과세금액 컬럼 '${options.taxableColumn}'`);
        }
      }
      
      if (missingColumns.length > 0) {
        console.error('Missing columns:', missingColumns);
        console.error('Available columns:', availableColumns);
        throw new Error(`다음 컬럼을 찾을 수 없습니다: ${missingColumns.join(', ')}`);
      }
    }
    
    return this.parseMonthlyData(data, options);
  }

  private processMergedCells(data: any[][], merges: any[]): any[][] {
    // Create a deep copy to avoid modifying original data
    const processedData = data.map(row => [...row]);
    
    // Process each merge
    merges.forEach(merge => {
      const value = processedData[merge.s.r]?.[merge.s.c];
      
      // Fill all cells in the merge range with the value
      for (let r = merge.s.r; r <= merge.e.r && r < processedData.length; r++) {
        for (let c = merge.s.c; c <= merge.e.c; c++) {
          if (!processedData[r]) {
            processedData[r] = [];
          }
          if (!processedData[r][c]) {
            processedData[r][c] = value;
          }
        }
      }
    });
    
    return processedData;
  }

  private combineMultiRowHeaders(data: any[][], headerRows: number[]): string[] {
    if (headerRows.length < 2) {
      // Single row header - return as is
      const headerRow = data[headerRows[0]] || [];
      return headerRow.map(cell => (cell || '').toString().trim()).filter(header => header);
    }
    
    const primaryRow = data[headerRows[0]] || [];
    const secondaryRow = data[headerRows[1]] || [];
    
    // Determine actual data range by finding the last meaningful column
    // Use manual implementation of findLastIndex for compatibility
    let maxPrimaryCol = -1;
    for (let i = primaryRow.length - 1; i >= 0; i--) {
      if (primaryRow[i] && primaryRow[i].toString().trim() !== '') {
        maxPrimaryCol = i;
        break;
      }
    }
    
    let maxSecondaryCol = -1;
    for (let i = secondaryRow.length - 1; i >= 0; i--) {
      if (secondaryRow[i] && secondaryRow[i].toString().trim() !== '') {
        maxSecondaryCol = i;
        break;
      }
    }
    
    const maxCols = Math.max(maxPrimaryCol, maxSecondaryCol) + 1;
    const combinedHeaders: string[] = [];
    
    console.log(`Header processing: maxPrimaryCol=${maxPrimaryCol}, maxSecondaryCol=${maxSecondaryCol}, maxCols=${maxCols}`);
    
    // Process each column individually to properly handle merged headers
    for (let colIndex = 0; colIndex < maxCols; colIndex++) {
      const primaryHeader = (primaryRow[colIndex] || '').toString().trim();
      const secondaryHeader = (secondaryRow[colIndex] || '').toString().trim();
      
      // Skip columns that have no data at all
      if (!primaryHeader && !secondaryHeader) {
        continue;
      }
      
      console.log(`Column ${colIndex}: primary="${primaryHeader}", secondary="${secondaryHeader}"`);
      
      let finalHeader = '';
      
      if (primaryHeader && secondaryHeader) {
        // Both headers exist - combine them
        finalHeader = `${primaryHeader} > ${secondaryHeader}`;
      } else if (primaryHeader && !secondaryHeader) {
        // Only primary header exists
        finalHeader = primaryHeader;
      } else if (!primaryHeader && secondaryHeader) {
        // Only secondary header exists - look for parent from merged cells
        let parentHeader = '';
        
        // Look left for the nearest non-empty primary header (for merged cells)
        for (let prevCol = colIndex - 1; prevCol >= 0; prevCol--) {
          const prevPrimaryHeader = (primaryRow[prevCol] || '').toString().trim();
          if (prevPrimaryHeader) {
            parentHeader = prevPrimaryHeader;
            break;
          }
        }
        
        if (parentHeader) {
          finalHeader = `${parentHeader} > ${secondaryHeader}`;
        } else {
          finalHeader = secondaryHeader;
        }
      }
      
      if (finalHeader) {
        combinedHeaders.push(finalHeader);
      }
    }
    
    // Remove duplicates to fix Cafe24-style merged header issues
    const uniqueHeaders = [];
    const seen = new Set();
    
    for (const header of combinedHeaders) {
      if (!seen.has(header)) {
        seen.add(header);
        uniqueHeaders.push(header);
      } else {
        console.log(`Removing duplicate header: "${header}"`);
      }
    }
    
    console.log('Combined multi-row headers:', uniqueHeaders);
    return uniqueHeaders;
  }

  private parseMonthlyData(rawData: any[], options: ProcessingOptions): MonthlyData[] {
    if (options.useTaxTypeClassification) {
      return this.parseMonthlyDataWithTaxType(rawData, options);
    }
    
    // Traditional method: separate columns for tax-exempt and taxable amounts
    const dateCol = options.dateColumn || '날짜';
    const taxExemptCol = options.taxExemptColumn || '면세금액';
    const taxableCol = options.taxableColumn || '과세금액';

    return rawData.map(row => {
      const date = this.parseDate(row[dateCol]);
      
      // Skip rows with invalid dates
      if (!date) {
        console.warn('Skipping row due to invalid date:', row[dateCol]);
        return null;
      }
      
      const taxExemptAmount = this.parseAmount(row[taxExemptCol]);
      const taxableAmount = this.parseAmount(row[taxableCol]);
      
      return {
        month: date.getMonth() + 1,
        year: date.getFullYear(),
        taxExemptAmount,
        taxableAmount
      };
    }).filter(item => 
      item !== null && 
      !isNaN(item.taxExemptAmount) && 
      !isNaN(item.taxableAmount)
    );
  }

  private parseMonthlyDataWithTaxType(rawData: any[], options: ProcessingOptions): MonthlyData[] {
    const dateCol = options.dateColumn || '날짜';
    const taxTypeCol = options.taxTypeColumn || '과세유형';
    
    // Default tax type values
    const taxExemptValues = options.taxExemptValues || ['면세', '면세상품', '0%', '영세율', 'FREE'];
    const taxableValues = options.taxableValues || ['과세', '과세상품', '10%', '부가세', 'TAX'];
    
    console.log(`Processing with tax type classification: ${taxTypeCol} column`);
    console.log(`Tax exempt values: ${taxExemptValues.join(', ')}`);
    console.log(`Taxable values: ${taxableValues.join(', ')}`);

    // Group data by date first, then aggregate by tax type
    const groupedByDate = new Map<string, {date: Date, taxExempt: number, taxable: number}>();

    rawData.forEach(row => {
      const date = this.parseDate(row[dateCol]);
      
      if (!date) {
        console.warn('Skipping row due to invalid date:', row[dateCol]);
        return;
      }
      
      const taxType = String(row[taxTypeCol] || '').trim();
      
      // Calculate total amount based on processing mode
      let amount = 0;
      if (options.useMultiColumnSum && options.amountColumns && options.amountColumns.length > 0) {
        // Sum multiple amount columns
        amount = options.amountColumns.reduce((sum, colName) => {
          const colValue = this.parseAmount(row[colName]);
          return sum + (isNaN(colValue) ? 0 : colValue);
        }, 0);
        
        console.log(`Multi-column sum for row: ${options.amountColumns.map(col => `${col}=${row[col]}`).join(', ')} = ${amount}`);
      } else {
        // Single amount column
        const amountCol = options.amountColumn || '금액';
        amount = this.parseAmount(row[amountCol]);
      }
      
      if (isNaN(amount) || amount === 0) {
        return; // Skip rows with invalid amounts
      }
      
      const dateKey = `${date.getFullYear()}-${date.getMonth() + 1}`;
      
      if (!groupedByDate.has(dateKey)) {
        groupedByDate.set(dateKey, {
          date,
          taxExempt: 0,
          taxable: 0
        });
      }
      
      const entry = groupedByDate.get(dateKey)!;
      
      // Classify based on tax type
      const isTaxExempt = taxExemptValues.some(val => 
        taxType.toLowerCase().includes(val.toLowerCase()) || 
        taxType === val
      );
      
      const isTaxable = taxableValues.some(val => 
        taxType.toLowerCase().includes(val.toLowerCase()) || 
        taxType === val
      );
      
      if (isTaxExempt) {
        entry.taxExempt += amount;
        console.log(`Added ${amount} to tax-exempt for ${dateKey} (type: ${taxType})`);
      } else if (isTaxable) {
        entry.taxable += amount;
        console.log(`Added ${amount} to taxable for ${dateKey} (type: ${taxType})`);
      } else {
        console.warn(`Unknown tax type: ${taxType}, treating as taxable`);
        entry.taxable += amount;
      }
    });

    // Convert grouped data to MonthlyData array
    return Array.from(groupedByDate.values()).map(entry => ({
      month: entry.date.getMonth() + 1,
      year: entry.date.getFullYear(),
      taxExemptAmount: entry.taxExempt,
      taxableAmount: entry.taxable
    }));
  }

  private parseDate(dateValue: any): Date | null {
    console.log(`Parsing date value: ${JSON.stringify(dateValue)} (type: ${typeof dateValue})`);
    
    // Handle null, undefined, or empty values
    if (dateValue == null || dateValue === '' || dateValue === undefined) {
      console.log('Date value is null/undefined/empty, returning null');
      return null;
    }
    
    // Already a Date object
    if (dateValue instanceof Date) {
      if (isNaN(dateValue.getTime())) {
        console.warn('Date object is invalid');
        return null;
      }
      console.log(`Valid Date object: ${dateValue.toISOString()}`);
      return dateValue;
    }
    
    // Handle Excel serial date (numbers)
    if (typeof dateValue === 'number') {
      if (dateValue < 1 || dateValue > 2958465) { // Valid Excel date range
        console.warn(`Number ${dateValue} outside valid Excel date range`);
        return null;
      }
      const excelDate = new Date((dateValue - 25569) * 86400 * 1000);
      console.log(`Excel serial date ${dateValue} → ${excelDate.toISOString()}`);
      return excelDate;
    }
    
    // Handle string date formats
    if (typeof dateValue === 'string') {
      const trimmed = dateValue.trim();
      if (trimmed === '') {
        console.log('Trimmed string is empty');
        return null;
      }
      
      console.log(`Attempting to parse string: "${trimmed}"`);
      
      // Strategy 0: Date range format (2024.10.01 ~ 2024.10.31)
      const dateRangeMatch = trimmed.match(/^(\d{4}\.\d{1,2}\.\d{1,2})\s*[~\-]\s*(\d{4}\.\d{1,2}\.\d{1,2})$/);
      if (dateRangeMatch) {
        const startDateStr = dateRangeMatch[1];
        const endDateStr = dateRangeMatch[2];
        
        // Parse start date (yyyy.MM.dd format)
        const startParts = startDateStr.split('.');
        if (startParts.length === 3) {
          const year = parseInt(startParts[0]);
          const month = parseInt(startParts[1]);
          const day = parseInt(startParts[2]);
          
          if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            const result = new Date(year, month - 1, day);
            console.log(`Parsed date range (using start date): ${trimmed} → ${result.toISOString()}`);
            return result;
          }
        }
      }
      
      // Strategy 0.1: Korean date range format (2024년 10월 1일 ~ 2024년 10월 31일)
      const koreanRangeMatch = trimmed.match(/^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일\s*[~\-]\s*(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일$/);
      if (koreanRangeMatch) {
        const year = parseInt(koreanRangeMatch[1]);
        const month = parseInt(koreanRangeMatch[2]);
        const day = parseInt(koreanRangeMatch[3]);
        
        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const result = new Date(year, month - 1, day);
          console.log(`Parsed Korean date range (using start date): ${trimmed} → ${result.toISOString()}`);
          return result;
        }
      }
      
      // Strategy 0.2: Month range format (2024.10 ~ 2024.11)
      const monthRangeMatch = trimmed.match(/^(\d{4}\.\d{1,2})\s*[~\-]\s*(\d{4}\.\d{1,2})$/);
      if (monthRangeMatch) {
        const startMonthStr = monthRangeMatch[1];
        const startParts = startMonthStr.split('.');
        
        if (startParts.length === 2) {
          const year = parseInt(startParts[0]);
          const month = parseInt(startParts[1]);
          
          if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12) {
            const result = new Date(year, month - 1, 1);
            console.log(`Parsed month range (using start month): ${trimmed} → ${result.toISOString()}`);
            return result;
          }
        }
      }
      
      // Strategy 0.3: ISO date range format (2024-10-01 ~ 2024-10-31)
      const isoDateRangeMatch = trimmed.match(/^(\d{4}-\d{1,2}-\d{1,2})\s*[~\-]\s*(\d{4}-\d{1,2}-\d{1,2})$/);
      if (isoDateRangeMatch) {
        const startDateStr = isoDateRangeMatch[1];
        const startDate = new Date(startDateStr);
        
        if (!isNaN(startDate.getTime()) && startDate.getFullYear() >= 1900 && startDate.getFullYear() <= 2100) {
          console.log(`Parsed ISO date range (using start date): ${trimmed} → ${startDate.toISOString()}`);
          return startDate;
        }
      }
      
      // Strategy 0.4: Mixed format range (2024.10.01~2024.10.31 without spaces)
      const compactRangeMatch = trimmed.match(/^(\d{4}\.\d{1,2}\.\d{1,2})[~\-](\d{4}\.\d{1,2}\.\d{1,2})$/);
      if (compactRangeMatch) {
        const startDateStr = compactRangeMatch[1];
        const startParts = startDateStr.split('.');
        
        if (startParts.length === 3) {
          const year = parseInt(startParts[0]);
          const month = parseInt(startParts[1]);
          const day = parseInt(startParts[2]);
          
          if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
            const result = new Date(year, month - 1, day);
            console.log(`Parsed compact date range: ${trimmed} → ${result.toISOString()}`);
            return result;
          }
        }
      }
      
      // Strategy 1: yyyy.MM format (2024.07)
      const yearMonthDotMatch = trimmed.match(/^(\d{4})\.(\d{1,2})$/);
      if (yearMonthDotMatch) {
        const year = parseInt(yearMonthDotMatch[1]);
        const month = parseInt(yearMonthDotMatch[2]);
        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12) {
          const result = new Date(year, month - 1, 1);
          console.log(`Parsed yyyy.MM format: ${trimmed} → ${result.toISOString()}`);
          return result;
        }
      }
      
      // Strategy 2: yyyy년 M월 format (2024년 7월)
      const koreanYearMonthMatch = trimmed.match(/^(\d{4})년\s*(\d{1,2})월$/);
      if (koreanYearMonthMatch) {
        const year = parseInt(koreanYearMonthMatch[1]);
        const month = parseInt(koreanYearMonthMatch[2]);
        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12) {
          const result = new Date(year, month - 1, 1);
          console.log(`Parsed Korean yyyy년 M월: ${trimmed} → ${result.toISOString()}`);
          return result;
        }
      }
      
      // Strategy 3: yyyy-MM-DD HH:mm:ss format (2024-07-21 09:30:00)
      const isoDateTimeMatch = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$/);
      if (isoDateTimeMatch) {
        const year = parseInt(isoDateTimeMatch[1]);
        const month = parseInt(isoDateTimeMatch[2]);
        const day = parseInt(isoDateTimeMatch[3]);
        const hour = parseInt(isoDateTimeMatch[4]);
        const minute = parseInt(isoDateTimeMatch[5]);
        const second = parseInt(isoDateTimeMatch[6]);
        
        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && 
            day >= 1 && day <= 31 && hour >= 0 && hour <= 23 && 
            minute >= 0 && minute <= 59 && second >= 0 && second <= 59) {
          const result = new Date(year, month - 1, day, hour, minute, second);
          console.log(`Parsed ISO datetime: ${trimmed} → ${result.toISOString()}`);
          return result;
        }
      }
      
      // Strategy 4: yyyy-MM-dd format (2024-07-21)
      const isoDateMatch = trimmed.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
      if (isoDateMatch) {
        const year = parseInt(isoDateMatch[1]);
        const month = parseInt(isoDateMatch[2]);
        const day = parseInt(isoDateMatch[3]);
        
        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const result = new Date(year, month - 1, day);
          console.log(`Parsed ISO date: ${trimmed} → ${result.toISOString()}`);
          return result;
        }
      }
      
      // Strategy 5: yyyy.MM.dd HH:mm:ss format (2024.07.21 09:30:00)
      const dotDateTimeMatch = trimmed.match(/^(\d{4})\.(\d{1,2})\.(\d{1,2})\s+(\d{1,2}):(\d{1,2}):(\d{1,2})$/);
      if (dotDateTimeMatch) {
        const year = parseInt(dotDateTimeMatch[1]);
        const month = parseInt(dotDateTimeMatch[2]);
        const day = parseInt(dotDateTimeMatch[3]);
        const hour = parseInt(dotDateTimeMatch[4]);
        const minute = parseInt(dotDateTimeMatch[5]);
        const second = parseInt(dotDateTimeMatch[6]);
        
        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && 
            day >= 1 && day <= 31 && hour >= 0 && hour <= 23 && 
            minute >= 0 && minute <= 59 && second >= 0 && second <= 59) {
          const result = new Date(year, month - 1, day, hour, minute, second);
          console.log(`Parsed dot datetime: ${trimmed} → ${result.toISOString()}`);
          return result;
        }
      }
      
      // Strategy 6: yyyy.MM.dd format (2024.07.21)
      const dotDateMatch = trimmed.match(/^(\d{4})\.(\d{1,2})\.(\d{1,2})$/);
      if (dotDateMatch) {
        const year = parseInt(dotDateMatch[1]);
        const month = parseInt(dotDateMatch[2]);
        const day = parseInt(dotDateMatch[3]);
        
        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const result = new Date(year, month - 1, day);
          console.log(`Parsed dot date: ${trimmed} → ${result.toISOString()}`);
          return result;
        }
      }
      
      // Strategy 7: Korean full date format (yyyy년 M월 D일)
      const koreanFullDateMatch = trimmed.match(/^(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일$/);
      if (koreanFullDateMatch) {
        const year = parseInt(koreanFullDateMatch[1]);
        const month = parseInt(koreanFullDateMatch[2]);
        const day = parseInt(koreanFullDateMatch[3]);
        
        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const result = new Date(year, month - 1, day);
          console.log(`Parsed Korean full date: ${trimmed} → ${result.toISOString()}`);
          return result;
        }
      }
      
      // Strategy 8: MM/DD/YYYY format
      const usFormatMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (usFormatMatch) {
        const month = parseInt(usFormatMatch[1]);
        const day = parseInt(usFormatMatch[2]);
        const year = parseInt(usFormatMatch[3]);
        
        if (year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const result = new Date(year, month - 1, day);
          console.log(`Parsed US format: ${trimmed} → ${result.toISOString()}`);
          return result;
        }
      }
      
      // Strategy 9: DD/MM/YYYY format
      const europeanFormatMatch = trimmed.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (europeanFormatMatch) {
        const day = parseInt(europeanFormatMatch[1]);
        const month = parseInt(europeanFormatMatch[2]);
        const year = parseInt(europeanFormatMatch[3]);
        
        // Only try this if the day > 12 (to avoid confusion with US format)
        if (day > 12 && year >= 1900 && year <= 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          const result = new Date(year, month - 1, day);
          console.log(`Parsed European format: ${trimmed} → ${result.toISOString()}`);
          return result;
        }
      }
      
      // Strategy 10: Try native Date constructor as fallback
      try {
        const nativeDate = new Date(trimmed);
        if (!isNaN(nativeDate.getTime()) && nativeDate.getFullYear() >= 1900 && nativeDate.getFullYear() <= 2100) {
          console.log(`Parsed using native Date constructor: ${trimmed} → ${nativeDate.toISOString()}`);
          return nativeDate;
        }
      } catch (error) {
        console.log(`Native Date constructor failed: ${error}`);
      }
      
      // Strategy 11: Try Date.parse() as final fallback
      try {
        const timestamp = Date.parse(trimmed);
        if (!isNaN(timestamp)) {
          const parsedDate = new Date(timestamp);
          if (parsedDate.getFullYear() >= 1900 && parsedDate.getFullYear() <= 2100) {
            console.log(`Parsed using Date.parse(): ${trimmed} → ${parsedDate.toISOString()}`);
            return parsedDate;
          }
        }
      } catch (error) {
        console.log(`Date.parse() failed: ${error}`);
      }
    }
    
    // If we get here, nothing worked
    console.warn(`Unable to parse date: ${JSON.stringify(dateValue)} (type: ${typeof dateValue})`);
    return null;
  }

  private parseAmount(value: any): number {
    if (typeof value === 'number') return value;
    if (typeof value === 'string') {
      // Remove currency symbols, commas, and spaces
      const cleaned = value.replace(/[₩$,\s]/g, '');
      return parseFloat(cleaned) || 0;
    }
    return 0;
  }

  /**
   * Automatically detect which row contains headers by analyzing data patterns
   */
  async detectHeaderRow(filePath: string, sheetName?: string): Promise<{
    headerRow: number;
    confidence: number;
    reasons: string[];
    isMultiRowHeader?: boolean;
    headerRows?: number[];
  }> {
    const workbook = XLSX.readFile(filePath);
    const worksheetName = sheetName || workbook.SheetNames[0];
    const worksheet = workbook.Sheets[worksheetName];
    
    if (!worksheet) {
      throw new Error(`Sheet ${worksheetName} not found in file`);
    }

    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
    
    if (data.length === 0) {
      throw new Error('No data found in Excel file');
    }

    const candidates: Array<{
      row: number;
      score: number;
      reasons: string[];
    }> = [];

    // Common header keywords in Korean and English
    const headerKeywords = [
      '날짜', '일자', 'date', 'Date', 'DATE',
      '금액', '면세', '과세', 'amount', 'Amount', 'AMOUNT',
      '총액', '합계', 'total', 'Total', 'TOTAL',
      '상품', '제품', 'product', 'Product', 'PRODUCT',
      '주문', '번호', 'order', 'Order', 'ORDER',
      '배송', '배송비', 'shipping', 'Shipping', 'SHIPPING',
      '구분', '분류', 'type', 'Type', 'TYPE',
      '매출', '수익', 'sales', 'Sales', 'SALES'
    ];

    // Analyze up to first 10 rows
    const maxRowsToCheck = Math.min(10, data.length);
    
    for (let rowIndex = 0; rowIndex < maxRowsToCheck; rowIndex++) {
      const row = data[rowIndex];
      if (!row || !Array.isArray(row)) continue;

      let score = 0;
      const reasons: string[] = [];

      // Skip empty rows
      const nonEmptyCells = row.filter(cell => cell != null && cell.toString().trim() !== '');
      if (nonEmptyCells.length === 0) continue;

      // 1. Check for high percentage of non-empty cells
      const nonEmptyRatio = nonEmptyCells.length / row.length;
      if (nonEmptyRatio > 0.5) {
        score += 20;
        reasons.push(`High non-empty ratio: ${(nonEmptyRatio * 100).toFixed(0)}%`);
      }

      // 2. Check for header keywords
      let keywordMatches = 0;
      nonEmptyCells.forEach(cell => {
        const cellStr = cell.toString().toLowerCase();
        if (headerKeywords.some(keyword => cellStr.includes(keyword.toLowerCase()))) {
          keywordMatches++;
        }
      });
      
      if (keywordMatches > 0) {
        score += keywordMatches * 15;
        reasons.push(`Found ${keywordMatches} header keywords`);
      }

      // 3. Check if most cells are text (not numbers)
      const textCells = nonEmptyCells.filter(cell => {
        const str = cell.toString();
        return isNaN(Number(str)) || str.match(/[가-힣a-zA-Z]/);
      });
      
      const textRatio = textCells.length / nonEmptyCells.length;
      if (textRatio > 0.7) {
        score += 25;
        reasons.push(`High text ratio: ${(textRatio * 100).toFixed(0)}%`);
      }

      // 4. Check if next row has different data types (header vs data pattern)
      if (rowIndex < data.length - 1) {
        const nextRow = data[rowIndex + 1];
        if (nextRow && Array.isArray(nextRow)) {
          const nextRowNumbers = nextRow.filter(cell => 
            cell != null && !isNaN(Number(cell)) && cell.toString().trim() !== ''
          ).length;
          
          const nextRowRatio = nextRowNumbers / nextRow.filter(c => c != null).length;
          if (textRatio > 0.7 && nextRowRatio > 0.5) {
            score += 30;
            reasons.push('Pattern change: text row followed by numeric row');
          }
        }
      }

      // 5. Bonus for being after empty or title rows
      if (rowIndex > 0) {
        const prevRow = data[rowIndex - 1];
        if (!prevRow || prevRow.filter(c => c != null && c.toString().trim() !== '').length === 0) {
          score += 10;
          reasons.push('Preceded by empty row');
        }
      }

      // 6. Check for common header patterns (no punctuation except allowed ones)
      const hasPunctuationPattern = nonEmptyCells.some(cell => {
        const str = cell.toString();
        return str.match(/[!@#$%^&*()+=\[\]{};':"\\|,.<>\/?]/);
      });
      
      if (!hasPunctuationPattern) {
        score += 5;
        reasons.push('No unusual punctuation');
      }

      if (score > 0) {
        candidates.push({ row: rowIndex, score, reasons });
      }
    }

    // Sort by score and return the best candidate
    candidates.sort((a, b) => b.score - a.score);
    
    if (candidates.length === 0) {
      // Default to row 0 if no good candidates found
      return {
        headerRow: 0,
        confidence: 30,
        reasons: ['No clear header pattern detected, defaulting to first row']
      };
    }

    const best = candidates[0];
    const confidence = Math.min(100, best.score);
    
    // Check for multi-row headers (like Cafe24 structure)
    const multiRowResult = this.detectMultiRowHeaders(data, best.row);
    
    return {
      headerRow: best.row,
      confidence,
      reasons: best.reasons,
      isMultiRowHeader: multiRowResult.isMultiRow,
      headerRows: multiRowResult.headerRows
    };
  }

  private detectMultiRowHeaders(data: any[][], primaryHeaderRow: number): {
    isMultiRow: boolean;
    headerRows: number[];
  } {
    // Check if the next row looks like a continuation of headers
    const nextRowIndex = primaryHeaderRow + 1;
    
    if (nextRowIndex >= data.length) {
      return { isMultiRow: false, headerRows: [primaryHeaderRow] };
    }
    
    const primaryRow = data[primaryHeaderRow] || [];
    const nextRow = data[nextRowIndex] || [];
    
    // Look for patterns indicating multi-row headers:
    // 1. Next row has empty cells aligned with primary row's content
    // 2. Next row has headers where primary row has empty cells
    // 3. Contains keywords like '과세금액', '면세금액' etc.
    
    let multiRowIndicators = 0;
    let totalCells = 0;
    
    const headerKeywords = ['과세금액', '면세금액', '금액', '수량', '건수', '합계'];
    
    for (let i = 0; i < Math.max(primaryRow.length, nextRow.length); i++) {
      totalCells++;
      
      const primaryCell = (primaryRow[i] || '').toString().trim();
      const nextCell = (nextRow[i] || '').toString().trim();
      
      // Case 1: Primary has content, next is empty (typical merged cell pattern)
      if (primaryCell && !nextCell) {
        multiRowIndicators++;
      }
      
      // Case 2: Primary is empty, next has header-like content
      if (!primaryCell && nextCell && headerKeywords.some(keyword => nextCell.includes(keyword))) {
        multiRowIndicators += 2; // Weight this higher
      }
      
      // Case 3: Both have content but next looks like sub-header
      if (primaryCell && nextCell && headerKeywords.some(keyword => nextCell.includes(keyword))) {
        multiRowIndicators++;
      }
    }
    
    // If more than 40% of cells suggest multi-row structure
    const multiRowScore = multiRowIndicators / totalCells;
    const isMultiRow = multiRowScore > 0.4;
    
    console.log(`Multi-row header detection: score=${multiRowScore}, isMultiRow=${isMultiRow}`);
    console.log(`Primary row (${primaryHeaderRow}):`, primaryRow);
    console.log(`Next row (${nextRowIndex}):`, nextRow);
    
    return {
      isMultiRow,
      headerRows: isMultiRow ? [primaryHeaderRow, nextRowIndex] : [primaryHeaderRow]
    };
  }

  async getColumnHeaders(filePath: string, options?: { headerRow?: number; autoDetect?: boolean }): Promise<string[]> {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    if (!worksheet) {
      throw new Error(`No sheets found in file`);
    }

    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (data.length === 0) {
      throw new Error('No data found in Excel file');
    }

    let headerDetection: any;
    
    // Use auto-detection if enabled or no headerRow specified
    if (options?.autoDetect !== false && options?.headerRow === undefined) {
      headerDetection = await this.detectHeaderRow(filePath, sheetName);
      console.log(`Auto-detected header row: ${headerDetection.headerRow} (confidence: ${headerDetection.confidence}%)`);
      console.log(`Reasons: ${headerDetection.reasons.join(', ')}`);
      
      if (headerDetection.isMultiRowHeader) {
        console.log(`Multi-row header detected: rows ${headerDetection.headerRows.join(', ')}`);
        // Use our improved combineMultiRowHeaders logic
        const combinedHeaders = this.combineMultiRowHeaders(data as any[][], headerDetection.headerRows);
        return combinedHeaders;
      }
    } else {
      headerDetection = { headerRow: options?.headerRow || 0 };
    }
    
    const headerRow = headerDetection.headerRow;
    
    if (headerRow >= data.length) {
      throw new Error(`Header row ${headerRow} exceeds data length ${data.length}`);
    }

    // Get the specified row as headers
    let headers = data[headerRow] as string[];
    
    // Handle merged cells in headers
    const merges = worksheet['!merges'];
    if (merges && merges.length > 0) {
      // Process each merge that affects the header row
      merges.forEach(merge => {
        if (merge.s.r <= headerRow && merge.e.r >= headerRow) { 
          // If merge includes the header row
          const sourceRow = merge.s.r;
          const sourceCol = merge.s.c;
          
          // Get the value from the top-left cell of the merge
          let value = null;
          if (sourceRow < data.length && data[sourceRow]) {
            const sourceRowData = data[sourceRow] as any[];
            value = sourceRowData[sourceCol];
          }
          
          // Fill the value across all merged columns in the header row
          for (let c = merge.s.c; c <= merge.e.c; c++) {
            if (!headers[c] || headers[c] === '') {
              headers[c] = value;
            }
          }
        }
      });
    }
    
    // Filter out any remaining null/empty headers
    return headers.filter(header => header != null && header.toString().trim() !== '');
  }

  /**
   * Get column headers with support for complex multi-row headers
   * This method handles scenarios where headers span multiple rows
   */
  async getMultiRowColumnHeaders(filePath: string, options?: { 
    headerRows?: number[], // Array of row indices to combine for headers
    separator?: string     // Separator for combining multi-row headers
  }): Promise<string[]> {
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    
    if (!worksheet) {
      throw new Error(`No sheets found in file`);
    }

    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    
    if (data.length === 0) {
      throw new Error('No data found in Excel file');
    }

    const headerRows = options?.headerRows || [0];
    const separator = options?.separator || ' - ';
    
    // Initialize combined headers array
    const maxCols = Math.max(...data.map(row => Array.isArray(row) ? row.length : 0));
    const combinedHeaders: string[] = new Array(maxCols).fill('');
    
    // Process merged cells first
    const merges = worksheet['!merges'];
    const processedMerges = new Map<string, any>();
    
    if (merges && merges.length > 0) {
      merges.forEach(merge => {
        // Store merge information for later use
        for (let r = merge.s.r; r <= merge.e.r; r++) {
          for (let c = merge.s.c; c <= merge.e.c; c++) {
            const key = `${r},${c}`;
            const mergeData = data[merge.s.r] as any[];
            processedMerges.set(key, {
              value: mergeData?.[merge.s.c],
              isTopLeft: r === merge.s.r && c === merge.s.c
            });
          }
        }
      });
    }
    
    // Build combined headers from specified rows
    headerRows.forEach(rowIndex => {
      if (rowIndex < data.length) {
        const row = data[rowIndex] as any[];
        
        for (let c = 0; c < maxCols; c++) {
          let cellValue = '';
          const mergeKey = `${rowIndex},${c}`;
          
          if (processedMerges.has(mergeKey)) {
            // Use value from merged cell
            const mergeInfo = processedMerges.get(mergeKey);
            cellValue = mergeInfo?.value || '';
          } else {
            // Use regular cell value
            cellValue = row?.[c] || '';
          }
          
          if (cellValue && cellValue.toString().trim()) {
            if (combinedHeaders[c]) {
              combinedHeaders[c] += separator + cellValue.toString().trim();
            } else {
              combinedHeaders[c] = cellValue.toString().trim();
            }
          }
        }
      }
    });
    
    // Filter out empty headers
    return combinedHeaders.filter(header => header.trim() !== '');
  }

  /**
   * Analyze worksheet structure including merged cells
   * Useful for debugging and understanding Excel file structure
   */
  async analyzeWorksheetStructure(filePath: string, sheetName?: string): Promise<{
    hasmergedCells: boolean,
    mergedCells: Array<{range: string, startCell: string, value: any}>,
    totalRows: number,
    totalColumns: number,
    headerRowSuggestion: number
  }> {
    const workbook = XLSX.readFile(filePath);
    const worksheetName = sheetName || workbook.SheetNames[0];
    const worksheet = workbook.Sheets[worksheetName];
    
    if (!worksheet) {
      throw new Error(`Sheet ${worksheetName} not found in file`);
    }

    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    const merges = worksheet['!merges'] || [];
    
    // Process merge information
    const mergedCells = merges.map(merge => {
      const startCell = XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
      const endCell = XLSX.utils.encode_cell({ r: merge.e.r, c: merge.e.c });
      const rowData = data[merge.s.r] as any[];
      const value = rowData?.[merge.s.c];
      
      return {
        range: `${startCell}:${endCell}`,
        startCell,
        value
      };
    });
    
    // Find the first row with substantial data (likely headers)
    let headerRowSuggestion = 0;
    for (let i = 0; i < Math.min(data.length, 10); i++) {
      const row = data[i] as any[];
      if (row && row.filter(cell => cell != null && cell.toString().trim() !== '').length > 2) {
        headerRowSuggestion = i;
        break;
      }
    }
    
    return {
      hasmergedCells: merges.length > 0,
      mergedCells,
      totalRows: data.length,
      totalColumns: Math.max(...data.map(row => Array.isArray(row) ? row.length : 0)),
      headerRowSuggestion
    };
  }

  /**
   * Comprehensive analysis of Excel file with header detection and structure insights
   */
  async analyzeExcelFile(filePath: string, sheetName?: string): Promise<{
    fileInfo: {
      name: string;
      path: string;
      size: number;
    };
    headerDetection: {
      detectedRow: number;
      confidence: number;
      reasons: string[];
      headers: string[];
    };
    structure: {
      totalRows: number;
      totalColumns: number;
      hasmergedCells: boolean;
      mergedCells: Array<{range: string, startCell: string, value: any}>;
    };
    dataPreview: {
      firstDataRow: number;
      sampleData: any[][];
      isEmpty: boolean;
    };
    recommendations: string[];
  }> {
    const stats = fs.statSync(filePath);
    const fileName = path.basename(filePath);
    
    // Detect headers
    const headerDetection = await this.detectHeaderRow(filePath, sheetName);
    const headers = await this.getColumnHeaders(filePath, { 
      headerRow: headerDetection.headerRow,
      autoDetect: false 
    });
    
    // Analyze structure
    const structure = await this.analyzeWorksheetStructure(filePath, sheetName);
    
    // Get data preview
    const workbook = XLSX.readFile(filePath);
    const worksheetName = sheetName || workbook.SheetNames[0];
    const worksheet = workbook.Sheets[worksheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
    
    const firstDataRow = headerDetection.headerRow + 1;
    const sampleData = data.slice(firstDataRow, Math.min(firstDataRow + 5, data.length));
    const isEmpty = sampleData.length === 0 || sampleData.every(row => 
      !row || row.every(cell => cell == null || cell.toString().trim() === '')
    );
    
    // Generate recommendations
    const recommendations: string[] = [];
    
    if (headerDetection.confidence < 80) {
      recommendations.push(`Low header detection confidence (${headerDetection.confidence}%). Consider manually specifying header row.`);
    }
    
    if (structure.hasmergedCells) {
      recommendations.push(`File contains ${structure.mergedCells.length} merged cells. Use processMergedCells for best results.`);
    }
    
    if (isEmpty) {
      recommendations.push('No data rows found after headers. Check if file contains actual data.');
    }
    
    const koreanHeaders = headers.filter(h => /[가-힣]/.test(h));
    if (koreanHeaders.length > 0) {
      recommendations.push(`File contains Korean headers: ${koreanHeaders.join(', ')}. Ensure proper encoding.`);
    }
    
    if (headers.some(h => h.includes('날짜') || h.toLowerCase().includes('date'))) {
      recommendations.push('Date columns detected. Use parseDate method for proper date handling.');
    }
    
    if (headers.some(h => h.includes('금액') || h.toLowerCase().includes('amount'))) {
      recommendations.push('Amount columns detected. Use parseAmount method for proper number formatting.');
    }
    
    return {
      fileInfo: {
        name: fileName,
        path: filePath,
        size: stats.size
      },
      headerDetection: {
        detectedRow: headerDetection.headerRow,
        confidence: headerDetection.confidence,
        reasons: headerDetection.reasons,
        headers
      },
      structure: {
        totalRows: structure.totalRows,
        totalColumns: structure.totalColumns,
        hasmergedCells: structure.hasmergedCells,
        mergedCells: structure.mergedCells
      },
      dataPreview: {
        firstDataRow,
        sampleData,
        isEmpty
      },
      recommendations
    };
  }

  async exportResults(results: any[], filePath: string, format: 'xlsx' | 'csv'): Promise<void> {
    if (format === 'xlsx') {
      const workbook = XLSX.utils.book_new();
      
      // Summary sheet
      const summaryData = results.map(r => ({
        '쇼핑몰': r.mallName,
        '연간 면세 합계': r.yearlyTotal.taxExempt,
        '연간 과세 합계': r.yearlyTotal.taxable,
        '연간 총 합계': r.yearlyTotal.total
      }));
      
      const summarySheet = XLSX.utils.json_to_sheet(summaryData);
      XLSX.utils.book_append_sheet(workbook, summarySheet, '요약');
      
      // Detail sheets for each mall
      results.forEach(result => {
        const detailData = result.monthlyTotals.map((m: any) => ({
          '년도': m.year,
          '월': m.month,
          '면세금액': m.taxExempt,
          '과세금액': m.taxable,
          '합계': m.total
        }));
        
        const detailSheet = XLSX.utils.json_to_sheet(detailData);
        XLSX.utils.book_append_sheet(workbook, detailSheet, result.mallName);
      });
      
      XLSX.writeFile(workbook, filePath);
    } else {
      // CSV export - summary only
      const csvData = results.map(r => 
        `${r.mallName},${r.yearlyTotal.taxExempt},${r.yearlyTotal.taxable},${r.yearlyTotal.total}`
      );
      
      const header = '쇼핑몰,연간 면세 합계,연간 과세 합계,연간 총 합계';
      fs.writeFileSync(filePath, [header, ...csvData].join('\n'), 'utf-8');
    }
  }
}