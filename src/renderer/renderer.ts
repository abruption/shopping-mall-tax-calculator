declare global {
  interface Window {
    electronAPI: {
      selectFiles: () => Promise<string[]>;
      processFiles: (filePaths: string[], options: any) => Promise<any>;
      exportResults: (results: any[], format: 'xlsx' | 'csv') => Promise<any>;
      getExcelColumns: (filePath: string) => Promise<any>;
    };
  }
}

class TaxCalculatorApp {
  private selectedFiles: string[] = [];
  private processedResults: any[] = [];

  constructor() {
    this.initializeEventListeners();
    this.initializeModeToggle();
    this.initializeMultiColumnToggle();
  }

  private initializeEventListeners() {
    document.getElementById('selectBtn')?.addEventListener('click', () => this.selectFiles());
    document.getElementById('processBtn')?.addEventListener('click', () => this.processFiles());
    document.getElementById('exportBtn')?.addEventListener('click', () => this.exportResults());
  }

  private initializeModeToggle() {
    const traditionalMode = document.getElementById('traditionalMode') as HTMLInputElement;
    const taxTypeMode = document.getElementById('taxTypeMode') as HTMLInputElement;
    
    traditionalMode?.addEventListener('change', () => this.toggleProcessingMode());
    taxTypeMode?.addEventListener('change', () => this.toggleProcessingMode());
  }

  private initializeMultiColumnToggle() {
    const multiColumnCheckbox = document.getElementById('useMultiColumnSum') as HTMLInputElement;
    multiColumnCheckbox?.addEventListener('change', () => this.toggleMultiColumnMode());
  }

  private toggleMultiColumnMode() {
    const multiColumnCheckbox = document.getElementById('useMultiColumnSum') as HTMLInputElement;
    const amountColumn = document.getElementById('amountColumn') as HTMLSelectElement;
    const amountColumnGroup = amountColumn?.closest('.option-group') as HTMLElement;
    const multiColumnConfig = document.getElementById('multiColumnConfig');
    
    if (multiColumnCheckbox?.checked) {
      if (amountColumnGroup) amountColumnGroup.style.display = 'none';
      multiColumnConfig!.style.display = 'block';
    } else {
      if (amountColumnGroup) amountColumnGroup.style.display = 'flex';
      multiColumnConfig!.style.display = 'none';
    }
  }

  private toggleProcessingMode() {
    const taxTypeMode = document.getElementById('taxTypeMode') as HTMLInputElement;
    const traditionalOptions = document.getElementById('traditionalOptions');
    const taxTypeOptions = document.getElementById('taxTypeOptions');
    const taxTypeConfig = document.getElementById('taxTypeConfig');
    
    if (taxTypeMode?.checked) {
      traditionalOptions!.style.display = 'none';
      taxTypeOptions!.style.display = 'grid';
      taxTypeConfig!.style.display = 'block';
    } else {
      traditionalOptions!.style.display = 'grid';
      taxTypeOptions!.style.display = 'none';
      taxTypeConfig!.style.display = 'none';
    }
  }


  private async selectFiles() {
    try {
      const files = await window.electronAPI.selectFiles();
      if (files && files.length > 0) {
        this.selectedFiles = files;
        this.showStatus(`${files.length}개 파일이 선택되었습니다.`);
        
        // Load column headers from the first file
        await this.loadColumnHeaders(files[0]);
        
        this.enableProcessButton();
      }
    } catch (error) {
      this.showError(`파일 선택 중 오류 발생: ${error.message}`);
    }
  }

  private async loadColumnHeaders(filePath: string) {
    try {
      const result = await window.electronAPI.getExcelColumns(filePath);
      
      if (result.success && result.columns) {
        this.updateColumnSelects(result.columns);
      } else {
        this.showError(`컬럼 정보를 읽을 수 없습니다: ${result.error}`);
      }
    } catch (error) {
      this.showError(`컬럼 정보 로드 중 오류: ${error.message}`);
    }
  }

  private updateColumnSelects(columns: string[]) {
    // Traditional mode selects
    const dateSelect = document.getElementById('dateColumn') as HTMLSelectElement;
    const taxExemptSelect = document.getElementById('taxExemptColumn') as HTMLSelectElement;
    const taxableSelect = document.getElementById('taxableColumn') as HTMLSelectElement;
    
    // Tax type mode selects
    const dateSelectTax = document.getElementById('dateColumnTax') as HTMLSelectElement;
    const taxTypeSelect = document.getElementById('taxTypeColumn') as HTMLSelectElement;
    const amountSelect = document.getElementById('amountColumn') as HTMLSelectElement;

    // All selects to update
    const allSelects = [dateSelect, taxExemptSelect, taxableSelect, dateSelectTax, taxTypeSelect, amountSelect];

    // Clear existing options
    allSelects.forEach(select => {
      if (select) {
        select.innerHTML = '<option value="">컬럼을 선택하세요</option>';
        select.disabled = false;
      }
    });

    // Add column options to all selects
    columns.forEach(column => {
      const option = new Option(column, column);
      allSelects.forEach(select => {
        if (select) {
          select.appendChild(option.cloneNode(true) as HTMLOptionElement);
        }
      });
    });

    // Update multi-column checkboxes
    this.updateMultiColumnCheckboxes(columns);

    // Try to auto-select based on common names
    this.autoSelectColumns(columns);
  }

  private updateMultiColumnCheckboxes(columns: string[]) {
    const container = document.getElementById('multiColumnSelectors');
    if (!container) return;

    container.innerHTML = '';

    // Filter columns that look like amount columns
    const amountColumns = columns.filter(col => {
      const lowerCol = col.toLowerCase();
      return (
        lowerCol.includes('금액') || lowerCol.includes('amount') ||
        lowerCol.includes('신용카드') || lowerCol.includes('현금') ||
        lowerCol.includes('기타') || lowerCol.includes('가격') ||
        lowerCol.includes('price') || lowerCol.includes('판매') ||
        lowerCol.includes('환불') || lowerCol.includes('리워드') ||
        /만원$/.test(col) || /원$/.test(col)
      ) && !lowerCol.includes('전체') && !lowerCol.includes('합계');
    });

    amountColumns.forEach(column => {
      const label = document.createElement('label');
      label.className = 'column-checkbox';
      
      const checkbox = document.createElement('input');
      checkbox.type = 'checkbox';
      checkbox.value = column;
      checkbox.name = 'amountColumns';
      
      // Auto-select columns that contain payment methods
      const autoSelectPatterns = ['신용카드', '현금', '기타', '판매'];
      if (autoSelectPatterns.some(pattern => column.includes(pattern))) {
        checkbox.checked = true;
        label.classList.add('selected');
      }
      
      checkbox.addEventListener('change', () => {
        if (checkbox.checked) {
          label.classList.add('selected');
        } else {
          label.classList.remove('selected');
        }
      });
      
      const span = document.createElement('span');
      span.textContent = column;
      
      label.appendChild(checkbox);
      label.appendChild(span);
      container.appendChild(label);
    });
  }

  private autoSelectColumns(columns: string[]) {
    // Traditional mode selects
    const dateSelect = document.getElementById('dateColumn') as HTMLSelectElement;
    const taxExemptSelect = document.getElementById('taxExemptColumn') as HTMLSelectElement;
    const taxableSelect = document.getElementById('taxableColumn') as HTMLSelectElement;
    
    // Tax type mode selects
    const dateSelectTax = document.getElementById('dateColumnTax') as HTMLSelectElement;
    const taxTypeSelect = document.getElementById('taxTypeColumn') as HTMLSelectElement;
    const amountSelect = document.getElementById('amountColumn') as HTMLSelectElement;

    // Common patterns for auto-selection (including multi-row header patterns)
    const datePatterns = ['날짜', '일자', 'date', 'Date', 'DATE', '거래일', '거래년월'];
    const taxExemptPatterns = ['면세', '면세금액', '면세액', 'tax-free', 'tax_free', 'exempted', '결제금액 > 면세금액'];
    const taxablePatterns = ['과세', '과세금액', '과세액', 'taxable', 'tax', 'taxed', '결제금액 > 과세금액'];
    const taxTypePatterns = ['과세유형', '과면세구분', '세금유형', '과세구분', 'tax_type', 'tax-type'];
    const amountPatterns = ['금액', '총금액', '가격', 'amount', 'price', 'total', '결제금액'];

    columns.forEach(column => {
      const lowerColumn = column.toLowerCase();
      
      // Auto-select for traditional mode
      if (datePatterns.some(pattern => column.includes(pattern) || lowerColumn.includes(pattern.toLowerCase()))) {
        if (dateSelect) dateSelect.value = column;
        if (dateSelectTax) dateSelectTax.value = column;
      }
      if (taxExemptPatterns.some(pattern => column.includes(pattern) || lowerColumn.includes(pattern.toLowerCase()))) {
        if (taxExemptSelect) taxExemptSelect.value = column;
      }
      if (taxablePatterns.some(pattern => column.includes(pattern) || lowerColumn.includes(pattern.toLowerCase()))) {
        if (taxableSelect) taxableSelect.value = column;
      }
      
      // Auto-select for tax type mode
      if (taxTypePatterns.some(pattern => column.includes(pattern) || lowerColumn.includes(pattern.toLowerCase()))) {
        if (taxTypeSelect) taxTypeSelect.value = column;
      }
      if (amountPatterns.some(pattern => column.includes(pattern) || lowerColumn.includes(pattern.toLowerCase()))) {
        if (amountSelect) amountSelect.value = column;
      }
    });
  }

  private async processFiles() {
    if (this.selectedFiles.length === 0) {
      this.showError('처리할 파일이 없습니다.');
      return;
    }

    const taxTypeMode = (document.getElementById('taxTypeMode') as HTMLInputElement).checked;
    let options: any;
    
    if (taxTypeMode) {
      // Tax type classification mode
      const dateColumn = (document.getElementById('dateColumnTax') as HTMLSelectElement).value;
      const taxTypeColumn = (document.getElementById('taxTypeColumn') as HTMLSelectElement).value;
      const amountColumn = (document.getElementById('amountColumn') as HTMLSelectElement).value;
      const sheetName = (document.getElementById('sheetNameTax') as HTMLInputElement).value || undefined;
      
      const taxExemptValues = (document.getElementById('taxExemptValues') as HTMLInputElement).value
        .split(',')
        .map(v => v.trim())
        .filter(v => v.length > 0);
      const taxableValues = (document.getElementById('taxableValues') as HTMLInputElement).value
        .split(',')
        .map(v => v.trim())
        .filter(v => v.length > 0);
      
      const useMultiColumnSum = (document.getElementById('useMultiColumnSum') as HTMLInputElement).checked;
      
      // Validate required fields based on mode
      if (!dateColumn || !taxTypeColumn) {
        this.showError('날짜 컬럼과 과세유형 컬럼을 선택해주세요.');
        return;
      }
      
      if (!useMultiColumnSum && !amountColumn) {
        this.showError('단일 컬럼 모드에서는 금액 컬럼을 선택해주세요.');
        return;
      }
      
      if (useMultiColumnSum) {
        const selectedColumns = Array.from(document.querySelectorAll('input[name="amountColumns"]:checked'))
          .map(checkbox => (checkbox as HTMLInputElement).value);
        
        console.log('Multi-column mode: selected columns =', selectedColumns);
        
        if (selectedColumns.length === 0) {
          this.showError('다중 컬럼 모드에서는 최소 하나의 금액 컬럼을 선택해주세요.');
          return;
        }
        
        options = {
          useTaxTypeClassification: true,
          useMultiColumnSum: true,
          dateColumn,
          taxTypeColumn,
          amountColumns: selectedColumns,
          taxExemptValues,
          taxableValues,
          sheetName
        };
      } else {
        options = {
          useTaxTypeClassification: true,
          useMultiColumnSum: false,
          dateColumn,
          taxTypeColumn,
          amountColumn,
          taxExemptValues,
          taxableValues,
          sheetName
        };
      }
    } else {
      // Traditional mode
      const dateColumn = (document.getElementById('dateColumn') as HTMLSelectElement).value;
      const taxExemptColumn = (document.getElementById('taxExemptColumn') as HTMLSelectElement).value;
      const taxableColumn = (document.getElementById('taxableColumn') as HTMLSelectElement).value;
      const sheetName = (document.getElementById('sheetName') as HTMLInputElement).value || undefined;
      
      if (!dateColumn || !taxExemptColumn || !taxableColumn) {
        this.showError('모든 컬럼을 선택해주세요.');
        return;
      }
      
      options = {
        useTaxTypeClassification: false,
        dateColumn,
        taxExemptColumn,
        taxableColumn,
        sheetName
      };
    }

    this.showLoading(true);
    this.clearError();

    try {
      const result = await window.electronAPI.processFiles(this.selectedFiles, options);
      
      if (result.success) {
        this.processedResults = result.data;
        this.displayResults(result.data);
        this.enableExportButton();
        this.showStatus('데이터 처리가 완료되었습니다.');
      } else {
        this.showError(`처리 중 오류 발생: ${result.error}`);
      }
    } catch (error) {
      this.showError(`처리 중 오류 발생: ${error.message}`);
    } finally {
      this.showLoading(false);
    }
  }

  private displayResults(results: any[]) {
    const resultsContainer = document.getElementById('results');
    if (!resultsContainer) return;

    resultsContainer.innerHTML = '';

    results.forEach(result => {
      const mallDiv = document.createElement('div');
      mallDiv.className = 'mall-result';
      
      mallDiv.innerHTML = `
        <h3>${result.mallName}</h3>
        <table>
          <thead>
            <tr>
              <th>년월</th>
              <th>면세금액</th>
              <th>과세금액</th>
              <th>합계</th>
            </tr>
          </thead>
          <tbody>
            ${result.monthlyTotals.map(month => `
              <tr>
                <td>${month.year}년 ${month.month}월</td>
                <td>${this.formatCurrency(month.taxExempt)}</td>
                <td>${this.formatCurrency(month.taxable)}</td>
                <td>${this.formatCurrency(month.total)}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
        <div class="summary">
          <h4>연간 합계</h4>
          <div class="summary-grid">
            <div class="summary-item">
              <div class="summary-label">면세금액</div>
              <div class="summary-value">${this.formatCurrency(result.yearlyTotal.taxExempt)}</div>
            </div>
            <div class="summary-item">
              <div class="summary-label">과세금액</div>
              <div class="summary-value">${this.formatCurrency(result.yearlyTotal.taxable)}</div>
            </div>
            <div class="summary-item">
              <div class="summary-label">총 합계</div>
              <div class="summary-value">${this.formatCurrency(result.yearlyTotal.total)}</div>
            </div>
          </div>
        </div>
      `;
      
      resultsContainer.appendChild(mallDiv);
    });
  }

  private formatCurrency(amount: number): string {
    return new Intl.NumberFormat('ko-KR', {
      style: 'currency',
      currency: 'KRW'
    }).format(amount);
  }

  private async exportResults() {
    if (this.processedResults.length === 0) {
      this.showError('내보낼 결과가 없습니다.');
      return;
    }

    try {
      const format = confirm('Excel 형식으로 내보내시겠습니까? (취소를 누르면 CSV로 내보냅니다)') ? 'xlsx' : 'csv';
      const result = await window.electronAPI.exportResults(this.processedResults, format);
      
      if (result.success) {
        this.showStatus(`파일이 저장되었습니다: ${result.path}`);
      }
    } catch (error) {
      this.showError(`내보내기 중 오류 발생: ${error.message}`);
    }
  }

  private enableProcessButton() {
    const btn = document.getElementById('processBtn') as HTMLButtonElement;
    if (btn) btn.disabled = false;
  }

  private enableExportButton() {
    const btn = document.getElementById('exportBtn') as HTMLButtonElement;
    if (btn) btn.disabled = false;
  }

  private showLoading(show: boolean) {
    const loading = document.getElementById('loading');
    if (loading) {
      loading.classList.toggle('active', show);
    }
  }

  private showError(message: string) {
    const errorElement = document.getElementById('errorMsg');
    if (errorElement) {
      errorElement.textContent = message;
      errorElement.classList.add('active');
    }
  }

  private clearError() {
    const errorElement = document.getElementById('errorMsg');
    if (errorElement) {
      errorElement.classList.remove('active');
    }
  }

  private showStatus(message: string) {
    const statusElement = document.getElementById('statusMsg');
    if (statusElement) {
      statusElement.textContent = message;
      statusElement.classList.add('active');
      setTimeout(() => {
        statusElement.classList.remove('active');
      }, 3000);
    }
  }
}

// Initialize app when DOM is ready
document.addEventListener('DOMContentLoaded', () => {
  new TaxCalculatorApp();
});