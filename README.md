# Shopping Mall Tax Calculator

A cross-platform Electron application for calculating monthly tax-exempt and taxable amounts from shopping mall Excel files with intelligent header detection and multi-format support.

## Features

- 📊 **Smart Excel Processing**: Automatic header detection and merged cell handling
- 💰 **Dual Processing Modes**: Traditional (separate tax columns) and Tax Type Classification
- 🔍 **Multi-Column Support**: Handle complex payment structures (Coupang-style multiple payment methods)
- 🧠 **Intelligent Header Detection**: Automatically detect and combine multi-row headers (Cafe24-style)
- 🌏 **Multi-language support** (Korean/English)
- 📁 **Flexible File Input**: Manual selection or batch processing
- 📤 **Export Results**: Excel or CSV format with detailed breakdown
- 🖥️ **Cross-platform**: Windows, macOS support
- 🛠️ **Advanced Configuration**: Customizable column mappings and tax type classifications

## Prerequisites

- Node.js 24 LTS, version 24.22.1 or later
- macOS 13 or later for the Electron 44 application
- pnpm 11.25.0 (the authoritative package manager)

## Installation

1. Clone the repository
2. Install dependencies:
   ```bash
   pnpm install --frozen-lockfile
   ```

## Supported Excel Formats

### Processing Modes

#### 1. Traditional Mode (Separate Tax Columns)
For files with separate tax-exempt and taxable amount columns:
- **Date Column**: `날짜`, `거래년월` (Date)
- **Tax Exempt Amount**: `면세금액` (Tax Exempt Amount)  
- **Taxable Amount**: `과세금액` (Taxable Amount)

#### 2. Tax Type Classification Mode
For files with single amount column and tax type classification:
- **Date Column**: `매출인식일`, `날짜` (Date)
- **Tax Type Column**: `과세유형` (Tax Classification)
- **Amount Column**: `금액`, `신용카드(판매)` (Amount)
- **Multi-Column Support**: Sum multiple payment methods (Coupang-style)

### Supported Mall Formats

- **Cafe24**: Multi-row merged headers (결제금액 > 과세금액/면세금액)
- **Coupang**: Tax type classification with multiple payment columns
- **Generic**: Standard separate column formats

### Intelligent Features

- **Auto Header Detection**: Automatically finds header rows in complex Excel files
- **Merged Cell Handling**: Properly processes merged headers and cells
- **Multi-Row Headers**: Combines parent/child header relationships
- **Flexible Column Mapping**: Customizable column names via UI

## Development

### Run in development mode:
```bash
pnpm start
```

### Run tests:
```bash
pnpm test
```

### Build the application:
```bash
pnpm build
```

## Building for Distribution

### Build for all platforms:
```bash
pnpm dist
```

### Build for specific platforms:
```bash
# macOS
pnpm dist:mac

# Windows
pnpm dist:win
```

Built applications will be available in the `build` directory.

## Usage

1. **Select Files**: Click "Excel 파일 선택" to choose your Excel files

2. **Choose Processing Mode**:
   - **기존 방식**: Traditional mode with separate tax columns
   - **과세유형 기준**: Tax type classification mode

3. **Configure Columns** (auto-detected, but customizable):
   - Date column selection
   - Tax/Amount column mappings
   - Multi-column sum options for complex payment structures

4. **Advanced Options**:
   - **Multi-Column Sum**: Enable for Coupang-style multiple payment methods
   - **Tax Type Configuration**: Set tax-exempt/taxable keywords
   - **Sheet Name**: Specify sheet or use auto-detection

5. **Process Data**: Click "데이터 처리" to calculate monthly and yearly totals

6. **Export Results**: Click "결과 내보내기" to save as Excel or CSV format

## Project Structure

```
tax-calculator-electron/
├── src/
│   ├── main/           # Main process files (IPC, file handling)
│   ├── renderer/       # Renderer process files (UI)
│   ├── services/       # Business logic services
│   │   ├── ExcelService.ts      # Excel processing with intelligent header detection
│   │   └── CalculationService.ts # Tax calculation logic
│   ├── types/          # TypeScript type definitions
│   └── __tests__/      # Unit tests and integration tests
├── dist/               # Compiled output (SWC)
├── build/              # Packaged applications
├── scripts/            # Build scripts
└── package.json
```

## Technologies Used

- **Electron** - Cross-platform desktop framework
- **TypeScript** - Type-safe JavaScript development  
- **SWC** - Fast TypeScript/JavaScript compiler (replaced Webpack)
- **xlsx (SheetJS)** - Excel file processing with advanced parsing
- **i18next** - Internationalization framework
- **Jest** - Testing framework with comprehensive test coverage
- **Electron Store** - Persistent configuration storage

## License

ISC
## Dependency security baseline

Use only `pnpm-lock.yaml`; do not generate a separate npm/yarn lockfile. SheetJS
CE 0.20.3 is pinned to the official CDN tarball with lockfile integrity because
npm's `xlsx` 0.18.5 is obsolete. See the [official installation guide](https://docs.sheetjs.com/docs/getting-started/installation/nodejs/).
Electron is pinned to 44.10.0 (Node 24.22.1 / Chromium 152.0.7977.169).
The Electron 37–44 migration retains context isolation and disabled Node integration;
the application uses BrowserWindow, dialog and IPC, not the removed desktop capture,
WebContents navigation or deprecated utility APIs. macOS 12 is no longer supported.
See [Electron breaking changes](https://www.electronjs.org/docs/latest/breaking-changes).

Validation: `pnpm type-check`, `pnpm test --runInBand`, `pnpm build`, and
`pnpm pack`. Audit development and shipped dependencies separately with
`pnpm audit` and `pnpm audit --prod`. Audit output is not proof of parser safety;
CDN packages and packaged Chromium/Node must also be checked against upstream advisories.
