const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Create test data with various header positions
const testFiles = [
  {
    name: 'headers-row-0.xlsx',
    data: [
      ['날짜', '면세금액', '과세금액', '총액'],
      ['2024-01-01', 100000, 50000, 150000],
      ['2024-01-02', 200000, 75000, 275000],
    ]
  },
  {
    name: 'headers-row-2.xlsx',
    data: [
      ['쇼핑몰 매출 데이터'],
      [],
      ['날짜', '면세금액', '과세금액', '총액'],
      ['2024-01-01', 100000, 50000, 150000],
      ['2024-01-02', 200000, 75000, 275000],
    ]
  },
  {
    name: 'merged-headers.xlsx',
    data: [
      ['매출 데이터', '', '금액', ''],
      ['날짜', '구분', '면세', '과세'],
      ['2024-01-01', '온라인', 100000, 50000],
      ['2024-01-02', '오프라인', 200000, 75000],
    ],
    merges: [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 1 } }, // "매출 데이터" across 2 cols
      { s: { r: 0, c: 2 }, e: { r: 0, c: 3 } }, // "금액" across 2 cols
    ]
  },
  {
    name: 'complex-headers.xlsx',
    data: [
      ['쇼핑몰 A - 2024년 매출 현황'],
      [],
      ['기간: 2024.01.01 ~ 2024.12.31'],
      [],
      ['날짜', '주문번호', '상품명', '면세금액', '과세금액', '배송비', '총액'],
      ['2024-01-01', 'ORD001', '상품A', 100000, 50000, 2500, 152500],
      ['2024-01-02', 'ORD002', '상품B', 200000, 75000, 0, 275000],
    ]
  }
];

// Create @xlsx directory structure
const xlsxDir = path.join(__dirname, '..', '@xlsx');
if (!fs.existsSync(xlsxDir)) {
  fs.mkdirSync(xlsxDir, { recursive: true });
}

// Create test Excel files
testFiles.forEach(file => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(file.data);
  
  // Add merges if specified
  if (file.merges) {
    ws['!merges'] = file.merges;
  }
  
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  
  const filePath = path.join(xlsxDir, file.name);
  XLSX.writeFile(wb, filePath);
  console.log(`Created: ${filePath}`);
});

console.log('Test Excel files created successfully!');