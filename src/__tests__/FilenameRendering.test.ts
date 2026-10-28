/** @jest-environment jsdom */
import * as fs from 'fs';
import * as path from 'path';
import * as ts from 'typescript';

// Execute the actual browser entrypoint as a script, as index.html does.
const script = ts.transpileModule(fs.readFileSync(path.join(__dirname, '../renderer/renderer.ts'), 'utf8'), {
  compilerOptions: { target: ts.ScriptTarget.ES2020, module: ts.ModuleKind.None }
}).outputText;

describe('untrusted filename rendering', () => {
  beforeEach(() => {
    document.body.innerHTML = '<div id="fileList"></div><div id="results"></div>';
  });

  it('renders markup-shaped filenames and mall names as literal text', () => {
    const app = new Function(`${script}; return new TaxCalculatorApp();`)();
    const name = '<img src=x onerror="window.filenameExecuted=true">.xlsx';
    app.updateFileStatus([name, 'C:\\imports\\normal.xlsx']);
    expect(document.querySelector('.file-item')!.textContent).toBe(name);
    expect(document.querySelectorAll('.file-item')).toHaveLength(2);
    expect(document.querySelector('#fileList img')).toBeNull();
    app.displayResults([{ mallName: name, monthlyTotals: [], yearlyTotal: { taxExempt: 0, taxable: 0, total: 0 } }]);
    expect(document.querySelector('#results h3')!.textContent).toBe(name);
    expect(document.querySelector('#results img')).toBeNull();
    expect((window as any).filenameExecuted).toBeUndefined();
    app.updateFileStatus(['replacement.xlsx']);
    expect(document.querySelectorAll('.file-item')).toHaveLength(1);
    app.displayResults([]);
    expect(document.querySelector('#results')!.childElementCount).toBe(0);
  });
});
