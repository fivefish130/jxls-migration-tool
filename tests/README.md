# JXLS Migration Tool Test Cases

This directory contains example files for testing the JXLS Migration Tool.

## 📁 Test File Descriptions

### 1. asin_top_tag.xls
**Description**: Single sheet test file
**Purpose**: Test basic JXLS 1.x → 2.x migration functionality
**Features**:
- Contains forEach loop instruction
- Simple data structure
- Suitable for verifying basic migration functionality

**Migration Command**:
```bash
# Run from project root
python jxls_migration_tool.py tests/asin_top_tag.xls -f -o tests/asin_top_tag_output.xlsx --keep-extension --verbose
```

### 2. tax_contract_multi_sheets_export.xlsx
**Description**: Multi-sheet test file
**Purpose**: Test complex multi-worksheet Excel file migration
**Features**:
- Contains 4 worksheets: SKU dimension, PO dimension, contract dimension, invoice details
- Each sheet has independent forEach loops
- Suitable for verifying multi-sheet processing capability

**Migration Command**:
```bash
# Run from project root
python jxls_migration_tool.py tests/tax_contract_multi_sheets_export.xlsx -f -o tests/tax_contract_output.xlsx --keep-extension --verbose
```

## 🛠️ Migration Tool Usage Guide

### Basic Syntax
```bash
python jxls_migration_tool.py <input file or directory> [options]
```

### Common Options

| Option | Description | Example |
|--------|-------------|---------|
| `-f` | Migrate single file (not directory) | `-f` |
| `-o, --output` | Specify output file/directory path | `-o output/` |
| `--keep-extension` | Keep original file extension | `--keep-extension` |
| `--dry-run` | Preview mode (doesn't modify files) | `--dry-run` |
| `--verbose` | Verbose logging output | `--verbose` |
| `-h, --help` | Show help information | `-h` |

### Complete Examples

#### Example 1: Single File Migration
```bash
python jxls_migration_tool.py tests/asin_top_tag.xls -f -o tests/output.xlsx --keep-extension --verbose
```

#### Example 2: Directory Batch Migration
```bash
# Migrate entire directory, keep extensions
python jxls_migration_tool.py exceltemplate_backup/ -o exceltemplate/ --keep-extension --verbose

# Migrate entire directory, convert to .xlsx
python jxls_migration_tool.py exceltemplate_backup/ -o exceltemplate/ --verbose
```

#### Example 3: Dry Run (Preview)
```bash
# Preview changes without actually modifying files
python jxls_migration_tool.py tests/ -f -o tests/output/ --dry-run --verbose
```

## ✅ Verify Migration Results

### View Output
After successful migration, the tool will display:
```
✅ Migration successful: output file path
🔧 Found X commands, converted Y commands
```

### Manual Verification
1. **Open output file** - Use Excel or WPS to open the generated `.xlsx` file
2. **Check comments** - Verify that A1 and A2 cells contain correct JXLS 2.x comments
3. **Validate expressions** - Confirm that `<jx:...>` tags have been converted to `${...}` expressions

### Automated Verification Script
You can use Python script to verify results:

```python
from openpyxl import load_workbook

# Open output file
wb = load_workbook('tests/output.xlsx')
ws = wb.active

# Check A1 comment
if ws['A1'].comment:
    print(f"A1 comment: {ws['A1'].comment.text}")

# Check A2 comment
if ws['A2'].comment:
    print(f"A2 comment: {ws['A2'].comment.text}")
```

## 🔍 Frequently Asked Questions

### Q: What if migration fails?
A: Use the `--verbose` option to view detailed error information, or check the log file.

### Q: What if comment positions are incorrect?
A: v3.4.1 has fixed the smart comment position issue. Comments now correctly appear in the first data cell of the data row.

### Q: What file formats are supported?
A: Supports `.xls` and `.xlsx` formats. The tool automatically detects the actual format, even if the extension doesn't match the actual format.

### Q: Can I batch migrate?
A: Yes. Simply specify the directory path, and the tool will automatically process all Excel files in the directory.

## 📋 Test Checklist

- [ ] Migration successful without errors
- [ ] JXLS instructions correctly converted
- [ ] A1 cell has jx:area comment
- [ ] A2 cell has jx:each comment
- [ ] Expression format correct `${...}`
- [ ] Original formatting preserved

## 📂 Test File Inventory

| File | Type | Sheets | JXLS Instructions | Status |
|------|------|--------|-------------------|--------|
| `asin_top_tag.xls` | .xls | 1 | forEach, area | ✅ Tested |
| `hot_sock_tag.xls` | .xls | 1 | forEach, area | ✅ Tested |
| `tax_contract_multi_sheets_export.xlsx` | .xlsx | 4 | forEach (per sheet) | ✅ Tested |

## 🐛 Reporting Issues

If you encounter issues during testing:
1. Use `--verbose` flag for detailed logs
2. Check the migration report in the output directory
3. Submit an issue with log file attached
4. Include the original Excel file for reproduction

## 📚 Additional Resources

- **Main Documentation**: See `/docs/USAGE.md` for detailed usage guide
- **API Reference**: See `/docs/API.md` for programmatic API documentation
- **Changelog**: See `/CHANGELOG.md` for version history

---

**Note**: This test suite is designed to verify the migration tool's functionality with real-world Excel templates. All test files are from actual production templates.
