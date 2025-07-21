import * as fs from 'fs';
import * as path from 'path';
import { ExcelService } from '../src/services/ExcelService';

interface ScanResult {
  file: string;
  detection: {
    headerRow: number;
    confidence: number;
    reasons: string[];
  };
  headers: string[];
  structure: {
    totalRows: number;
    totalColumns: number;
    hasMergedCells: boolean;
    mergedCellCount?: number;
  };
  error?: string;
}

async function scanExcelHeaders() {
  const excelService = new ExcelService();
  const xlsxDir = path.join(__dirname, '..', '@xlsx');
  
  console.log('Scanning Excel files in:', xlsxDir);
  console.log('=' .repeat(80));
  
  // Find all Excel files recursively
  const excelFiles: string[] = [];
  
  function findExcelFiles(dir: string) {
    const files = fs.readdirSync(dir);
    
    for (const file of files) {
      const fullPath = path.join(dir, file);
      const stat = fs.statSync(fullPath);
      
      if (stat.isDirectory()) {
        findExcelFiles(fullPath);
      } else if (file.match(/\.(xlsx|xls)$/i)) {
        excelFiles.push(fullPath);
      }
    }
  }
  
  if (fs.existsSync(xlsxDir)) {
    findExcelFiles(xlsxDir);
  } else {
    console.error(`Directory not found: ${xlsxDir}`);
    return;
  }
  
  console.log(`Found ${excelFiles.length} Excel file(s)\n`);
  
  const results: ScanResult[] = [];
  
  for (const filePath of excelFiles) {
    const relativePath = path.relative(xlsxDir, filePath);
    console.log(`\nAnalyzing: ${relativePath}`);
    console.log('-'.repeat(60));
    
    try {
      // Detect header row
      const detection = await excelService.detectHeaderRow(filePath);
      
      // Get headers using detected row
      const headers = await excelService.getColumnHeaders(filePath, { 
        headerRow: detection.headerRow,
        autoDetect: false // We already detected
      });
      
      // Analyze structure
      const structure = await excelService.analyzeWorksheetStructure(filePath);
      
      const result: ScanResult = {
        file: relativePath,
        detection,
        headers,
        structure: {
          totalRows: structure.totalRows,
          totalColumns: structure.totalColumns,
          hasMergedCells: structure.hasmergedCells,
          mergedCellCount: structure.mergedCells.length
        }
      };
      
      results.push(result);
      
      // Print results for this file
      console.log(`Header Detection:`);
      console.log(`  - Row: ${detection.headerRow}`);
      console.log(`  - Confidence: ${detection.confidence}%`);
      console.log(`  - Reasons: ${detection.reasons.join('; ')}`);
      
      console.log(`\nHeaders Found (${headers.length}):`);
      headers.forEach((header, idx) => {
        console.log(`  ${idx + 1}. "${header}"`);
      });
      
      console.log(`\nStructure:`);
      console.log(`  - Total Rows: ${structure.totalRows}`);
      console.log(`  - Total Columns: ${structure.totalColumns}`);
      console.log(`  - Has Merged Cells: ${structure.hasmergedCells}`);
      if (structure.hasmergedCells) {
        console.log(`  - Merged Cell Count: ${structure.mergedCells.length}`);
        console.log(`  - Merged Ranges:`);
        structure.mergedCells.slice(0, 5).forEach(merge => {
          console.log(`    - ${merge.range}: "${merge.value}"`);
        });
        if (structure.mergedCells.length > 5) {
          console.log(`    ... and ${structure.mergedCells.length - 5} more`);
        }
      }
      
    } catch (error) {
      const result: ScanResult = {
        file: relativePath,
        detection: { headerRow: -1, confidence: 0, reasons: [] },
        headers: [],
        structure: {
          totalRows: 0,
          totalColumns: 0,
          hasMergedCells: false
        },
        error: error instanceof Error ? error.message : String(error)
      };
      
      results.push(result);
      console.log(`ERROR: ${result.error}`);
    }
  }
  
  // Generate summary report
  console.log('\n' + '='.repeat(80));
  console.log('SUMMARY REPORT');
  console.log('='.repeat(80));
  
  console.log(`\nTotal Files Scanned: ${results.length}`);
  console.log(`Successful: ${results.filter(r => !r.error).length}`);
  console.log(`Failed: ${results.filter(r => r.error).length}`);
  
  // Group by confidence levels
  const byConfidence = {
    high: results.filter(r => !r.error && r.detection.confidence >= 80),
    medium: results.filter(r => !r.error && r.detection.confidence >= 50 && r.detection.confidence < 80),
    low: results.filter(r => !r.error && r.detection.confidence < 50)
  };
  
  console.log(`\nConfidence Distribution:`);
  console.log(`  - High (80-100%): ${byConfidence.high.length} files`);
  console.log(`  - Medium (50-79%): ${byConfidence.medium.length} files`);
  console.log(`  - Low (0-49%): ${byConfidence.low.length} files`);
  
  // Common header patterns
  const headerPatterns = new Map<string, number>();
  results.filter(r => !r.error).forEach(r => {
    r.headers.forEach(h => {
      headerPatterns.set(h, (headerPatterns.get(h) || 0) + 1);
    });
  });
  
  console.log(`\nMost Common Headers:`);
  const sortedHeaders = Array.from(headerPatterns.entries())
    .sort((a, b) => b[1] - a[1])
    .slice(0, 10);
  
  sortedHeaders.forEach(([header, count]) => {
    console.log(`  - "${header}": found in ${count} file(s)`);
  });
  
  // Save detailed report to JSON
  const reportPath = path.join(__dirname, '..', 'excel-header-scan-report.json');
  fs.writeFileSync(reportPath, JSON.stringify(results, null, 2));
  console.log(`\nDetailed report saved to: ${reportPath}`);
}

// Run the scanner
scanExcelHeaders().catch(console.error);