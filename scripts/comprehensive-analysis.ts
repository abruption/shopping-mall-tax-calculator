import { ExcelService } from '../src/services/ExcelService';
import * as path from 'path';
import * as fs from 'fs';

async function comprehensiveAnalysis() {
  const excelService = new ExcelService();
  const xlsxDir = path.join(__dirname, '..', '@xlsx');
  
  console.log('='.repeat(80));
  console.log('COMPREHENSIVE EXCEL FILE ANALYSIS');
  console.log('='.repeat(80));
  
  // Find all Excel files
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
  }
  
  for (const filePath of excelFiles) {
    const relativePath = path.relative(xlsxDir, filePath);
    
    console.log(`\n${'='.repeat(60)}`);
    console.log(`Analyzing: ${relativePath}`);
    console.log('='.repeat(60));
    
    try {
      const analysis = await excelService.analyzeExcelFile(filePath);
      
      // File Info
      console.log('\n📄 FILE INFORMATION:');
      console.log(`   Name: ${analysis.fileInfo.name}`);
      console.log(`   Size: ${(analysis.fileInfo.size / 1024).toFixed(2)} KB`);
      
      // Header Detection
      console.log('\n🔍 HEADER DETECTION:');
      console.log(`   Detected Row: ${analysis.headerDetection.detectedRow}`);
      console.log(`   Confidence: ${analysis.headerDetection.confidence}%`);
      console.log(`   Headers Found: ${analysis.headerDetection.headers.length}`);
      analysis.headerDetection.headers.forEach((header, idx) => {
        console.log(`     ${idx + 1}. "${header}"`);
      });
      console.log(`   Detection Reasons:`);
      analysis.headerDetection.reasons.forEach(reason => {
        console.log(`     • ${reason}`);
      });
      
      // Structure
      console.log('\n🏗️  STRUCTURE:');
      console.log(`   Total Rows: ${analysis.structure.totalRows}`);
      console.log(`   Total Columns: ${analysis.structure.totalColumns}`);
      console.log(`   Has Merged Cells: ${analysis.structure.hasmergedCells}`);
      if (analysis.structure.hasmergedCells) {
        console.log(`   Merged Cells: ${analysis.structure.mergedCells.length}`);
        analysis.structure.mergedCells.slice(0, 3).forEach(merge => {
          console.log(`     • ${merge.range}: "${merge.value}"`);
        });
        if (analysis.structure.mergedCells.length > 3) {
          console.log(`     ... and ${analysis.structure.mergedCells.length - 3} more`);
        }
      }
      
      // Data Preview
      console.log('\n📊 DATA PREVIEW:');
      console.log(`   First Data Row: ${analysis.dataPreview.firstDataRow}`);
      console.log(`   Is Empty: ${analysis.dataPreview.isEmpty}`);
      if (!analysis.dataPreview.isEmpty && analysis.dataPreview.sampleData.length > 0) {
        console.log(`   Sample Data (first ${Math.min(3, analysis.dataPreview.sampleData.length)} rows):`);
        analysis.dataPreview.sampleData.slice(0, 3).forEach((row, idx) => {
          const displayRow = row.slice(0, 4).map(cell => 
            cell != null ? cell.toString().substring(0, 15) : 'null'
          ).join(' | ');
          console.log(`     Row ${idx + 1}: ${displayRow}${row.length > 4 ? ' | ...' : ''}`);
        });
      }
      
      // Recommendations
      console.log('\n💡 RECOMMENDATIONS:');
      if (analysis.recommendations.length === 0) {
        console.log('   ✅ No specific recommendations - file looks good!');
      } else {
        analysis.recommendations.forEach(rec => {
          console.log(`   • ${rec}`);
        });
      }
      
    } catch (error) {
      console.log(`\n❌ ERROR: ${error instanceof Error ? error.message : error}`);
    }
  }
  
  console.log('\n' + '='.repeat(80));
  console.log('ANALYSIS COMPLETE');
  console.log('='.repeat(80));
}

// Run the comprehensive analysis
comprehensiveAnalysis().catch(console.error);