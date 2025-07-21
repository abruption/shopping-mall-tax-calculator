import { ExcelService } from '../src/services/ExcelService';
import * as path from 'path';
import * as fs from 'fs';

async function demonstrateHeaderDetection() {
  const excelService = new ExcelService();
  const xlsxDir = path.join(__dirname, '..', '@xlsx');
  
  console.log('='.repeat(80));
  console.log('DYNAMIC HEADER DETECTION DEMONSTRATION');
  console.log('='.repeat(80));
  
  const testFiles = [
    'headers-row-0.xlsx',
    'headers-row-2.xlsx', 
    'merged-headers.xlsx',
    'complex-headers.xlsx'
  ];
  
  for (const filename of testFiles) {
    const filePath = path.join(xlsxDir, filename);
    
    if (!fs.existsSync(filePath)) {
      console.log(`\nSkipping ${filename} - file not found`);
      continue;
    }
    
    console.log(`\n${'='.repeat(60)}`);
    console.log(`File: ${filename}`);
    console.log('='.repeat(60));
    
    try {
      // 1. Detect header row
      console.log('\n1. HEADER DETECTION:');
      const detection = await excelService.detectHeaderRow(filePath);
      console.log(`   - Detected Row: ${detection.headerRow}`);
      console.log(`   - Confidence: ${detection.confidence}%`);
      console.log(`   - Analysis:`);
      detection.reasons.forEach(reason => {
        console.log(`     • ${reason}`);
      });
      
      // 2. Get headers using auto-detection
      console.log('\n2. HEADERS (Auto-detected):');
      const autoHeaders = await excelService.getColumnHeaders(filePath);
      autoHeaders.forEach((header, idx) => {
        console.log(`   ${idx + 1}. "${header}"`);
      });
      
      // 3. Analyze worksheet structure
      console.log('\n3. WORKSHEET STRUCTURE:');
      const structure = await excelService.analyzeWorksheetStructure(filePath);
      console.log(`   - Total Rows: ${structure.totalRows}`);
      console.log(`   - Total Columns: ${structure.totalColumns}`);
      console.log(`   - Has Merged Cells: ${structure.hasmergedCells}`);
      
      if (structure.hasmergedCells) {
        console.log(`   - Merged Cells:`);
        structure.mergedCells.slice(0, 3).forEach(merge => {
          console.log(`     • ${merge.range}: "${merge.value}"`);
        });
      }
      
      // 4. Try to read data
      console.log('\n4. DATA PARSING (First 3 rows):');
      try {
        const data = await excelService.readExcelFile(filePath, {
          dateColumn: '날짜',
          taxExemptColumn: '면세금액',
          taxableColumn: '과세금액'
        });
        
        data.slice(0, 3).forEach((row, idx) => {
          console.log(`   Row ${idx + 1}: ${row.year}/${String(row.month).padStart(2, '0')} - 면세: ${row.taxExemptAmount.toLocaleString()}, 과세: ${row.taxableAmount.toLocaleString()}`);
        });
      } catch (parseError) {
        console.log(`   Error parsing data: ${parseError instanceof Error ? parseError.message : parseError}`);
      }
      
    } catch (error) {
      console.log(`\nERROR: ${error instanceof Error ? error.message : error}`);
    }
  }
  
  // Multi-row header example
  console.log(`\n${'='.repeat(60)}`);
  console.log('MULTI-ROW HEADER EXAMPLE');
  console.log('='.repeat(60));
  
  const mergedFile = path.join(xlsxDir, 'merged-headers.xlsx');
  if (fs.existsSync(mergedFile)) {
    try {
      const multiHeaders = await excelService.getMultiRowColumnHeaders(mergedFile, {
        headerRows: [0, 1],
        separator: ' > '
      });
      
      console.log('\nCombined Headers (Row 0 + Row 1):');
      multiHeaders.forEach((header, idx) => {
        console.log(`   ${idx + 1}. "${header}"`);
      });
    } catch (error) {
      console.log(`Error: ${error instanceof Error ? error.message : error}`);
    }
  }
  
  console.log('\n' + '='.repeat(80));
  console.log('DEMO COMPLETE');
  console.log('='.repeat(80));
}

// Run the demonstration
demonstrateHeaderDetection().catch(console.error);