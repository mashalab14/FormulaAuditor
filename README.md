# Formula Auditor - Google Sheets Add-on

A modular Google Sheets Add-on for formatting formula cells with custom styles.

## 📁 Project Structure

```
FormulaAuditor/
├── Code.gs                  # Main orchestrator
├── validation.js            # Input validation functions
├── sheetContext.js          # Sheet and range context setup
├── formulaDetection.js      # Formula cell detection
├── styleFetcher.js          # Current style retrieval
├── styleTransformer.js      # Style transformation in memory
├── styleWriter.js           # Bulk style application
├── debugLogger.js           # Debug information logging
├── cancellation.js          # Cancellation handling
├── formatFormulaCells.html  # Sidebar UI
└── appsscript.json          # Apps Script manifest
```

## 🏗️ Architecture

### Modular Design Pattern

The add-on follows a **strictly modular architecture** with single-responsibility functions:

#### **Main Orchestrator** (`Code.gs`)
Coordinates the formatting workflow in 5 steps:
1. **Validation** → `validation.js`
2. **Setup** → `sheetContext.js`, `debugLogger.js`
3. **Detection** → `formulaDetection.js`
4. **Transformation** → `styleTransformer.js`
5. **Execution** → `styleWriter.js`

#### **Module Files**

| Module | Responsibility | Key Function |
|--------|---------------|--------------|
| `validation.js` | Validates input styles | `isValidInput(styles)` |
| `sheetContext.js` | Sets up sheet context | `getSheetContext()` |
| `formulaDetection.js` | Finds formula cells | `findFormulaCells(range)` |
| `styleFetcher.js` | Fetches current styles | `fetchCurrentStyles(range)` |
| `styleTransformer.js` | Applies styles to memory | `applyStylesToMemory(...)` |
| `styleWriter.js` | Writes to sheet | `writeStylesToSheet(...)` |
| `debugLogger.js` | Logs debug info | `logDebugInfo(context)` |
| `cancellation.js` | Handles cancellation | `cancelFormatting()` |

## ✨ Features

- ✅ **Batch Processing**: Handles 1000+ cells efficiently
- ✅ **Cancellation Support**: Users can cancel long operations
- ✅ **Debug Logging**: Comprehensive console logging
- ✅ **Formula Detection**: Robust detection using `.toString().startsWith("=")`
- ✅ **Custom Formatting**: Bold, Italic, Underline, Strikethrough, Colors
- ✅ **Modern UI**: Clean sidebar with tooltips and live preview

## 🚀 Usage

1. Open your Google Sheet
2. Click **Formula Auditor** → **🎨 Format Formula Cells**
3. Select formatting options in the sidebar
4. Click **Apply** to format all formula cells
5. Use **Cancel** to stop ongoing operations

## 🛠️ Development

### Prerequisites
- [clasp](https://github.com/google/clasp) - Google Apps Script CLI
- Node.js and npm

### Deploy to Apps Script

```bash
# Push all files to Google Apps Script
clasp push
```

### Push to GitHub

```bash
git add .
git commit -m "Your commit message"
git push origin main
```

## 📋 Technical Details

- **Batch Size**: 1000 cells per flush
- **Cache Duration**: 10 minutes for cancellation flag
- **Range Strategy**: Uses full grid range for safety
- **Style Arrays**: In-memory transformation before bulk write

## 📄 License

MIT License - Feel free to use and modify!

## 👤 Author

Created with ❤️ for efficient Google Sheets formula management
