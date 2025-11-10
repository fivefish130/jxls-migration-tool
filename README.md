# JXLS Migration Tool


[![Version](https://img.shields.io/badge/version-3.0-blue.svg)](https://github.com/your-org/jxls-migration-tool)
[![Python](https://img.shields.io/badge/python-3.6+-green.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

A powerful, production-ready tool for automated migration from JXLS 1.x to JXLS 2.14.0 Excel templates.
**中文文档**: [中文版 README](docs/README_ZH.md) | [中文使用指南](docs/USAGE_ZH.md)



## ✨ Features

- **Complete JXLS Instruction Support**: Migrates all JXLS instructions (forEach, if, out, area, multiSheet)
- **Smart File Format Detection**: Auto-detects Excel file format regardless of extension
- **Format Preservation**: Maintains all cell styles, column widths, row heights, merged cells
- **Windows Terminal Optimized**: Auto-detects and configures UTF-8 support
- **Detailed Reporting**: Generates Markdown, JSON, and DEBUG logs
- **Production Ready**: Comprehensive error handling and logging

## 📦 Installation

### Requirements
- Python 3.6+
- xlrd 2.0.1
- openpyxl

### Quick Install
```bash
pip install xlrd==2.0.1 openpyxl
```

## 🚀 Quick Start

### Migrate a directory
```bash
# Keep original file extensions but convert to .xlsx format
# .xls files will have .xls extension but .xlsx content (Jxls 2.14.0 compatible)
# .xlsx files will remain .xlsx extension
python jxls_migration_tool.py input_directory --keep-extension

# Specify output directory
python jxls_migration_tool.py input_directory -o output_directory --keep-extension

# Dry run (preview changes without modifying files)
python jxls_migration_tool.py input_directory --dry-run --verbose
```

**Note**: `--keep-extension` preserves the original file extension but ensures all files have .xlsx content for Jxls 2.14.0 compatibility. This avoids the need to modify backend code.

### Migrate a single file
```bash
python jxls_migration_tool.py input.xls -f output.xlsx
```

## 📋 JXLS Instruction Conversion

| JXLS 1.x | JXLS 2.14.0 | Description |
|----------|-------------|-------------|
| `<jx:forEach items="..." var="...">` | `jx:each(items="..." var="..." lastCell="...")` | Comment-based, auto-removes tag rows |
| `<jx:if test="...">` | `jx:if(condition="..." lastCell="...")` | test → condition, comment-based |
| `<jx:out select="..."/>` | `${...}` | Direct expression replacement |
| `<jx:area lastCell="...">` | `jx:area(lastCell="...")` | Preserve existing or auto-generate |
| `<jx:multiSheet data="...">` | `jx:multiSheet(data="...")` | Full multi-sheet support |

## 📊 Real-World Results

**Migration Statistics (923 Excel files)**:
- ✅ Successfully migrated: 50 JXLS templates
- ⏭️  Skipped: 873 files (no JXLS instructions)
- ❌ Failed: 0 files
- 🎯 Success rate: **100%**

**Module Distribution**:
- Module A: 7 templates
- Module B: 1 template
- Module C: 42 templates

**Command Statistics**:
- Total JXLS commands found: 106
- Successfully converted: 106
- Main types: forEach (50), area (50), if (6)

## 📖 Documentation

- **[Usage Guide](docs/USAGE.md)** - Detailed usage instructions
- **[使用指南](docs/USAGE_ZH.md)** - 中文版详细使用说明
- **[API Documentation](docs/API.md)** - Programmatic API reference
- **[Examples](examples/)** - Code examples and use cases
- **[Changelog](docs/CHANGELOG.md)** - Version history and changes
```
usage: jxls_migration_tool.py [-h] [-o OUTPUT] [-f] [--keep-extension]
                              [--dry-run] [-v] [--verbose]
                              input_path

JXLS 1.x → 2.14.0 Migration Tool

positional arguments:
  input_path              Input file or directory path

optional arguments:
  -h, --help              show this help message and exit
  -o OUTPUT, --output OUTPUT
                          Output directory (default: same as input)
  -f OUTPUT_FILE, --file OUTPUT_FILE
                          Output file (for single file migration)
  --keep-extension        Keep original file extensions
  --dry-run               Preview changes without modifying files
  -v, --verbose           Enable verbose logging
```

## 🎯 Migration Example

### Before (JXLS 1.x)
```
Row 1: <jx:area lastCell="E5">
Row 2: Purchase Order List
Row 3: <jx:forEach items="datas" var="item">
Row 4: ${item.sku} | ${item.qty} | ${item.price}
Row 5: </jx:forEach>
Row 6: Total: ${total}
```

### After (JXLS 2.x)
```
Row 1: (blank)
       └─ [Comment] jx:area(lastCell="E4")
Row 2: Purchase Order List
Row 3: ${item.sku} | ${item.qty} | ${item.price}
       └─ [Comment] jx:each(items="datas" var="item" lastCell="C3")
Row 4: Total: ${total}
```

**Key Changes**:
- ✅ Deleted tag rows (3rd and 5th)
- ✅ Added jx:area comment at row 1
- ✅ Added jx:each comment at row 3
- ✅ Auto-calculated lastCell="C3" (column C, row 3)

## 🔧 Smart File Format Detection

The tool auto-detects the actual Excel file format by reading file headers, not relying on extensions:

| File Header | Format | Description |
|------------|--------|-------------|
| `D0 CF 11 E0 A1 B1 1A E1` | XLS | OLE2/Compound Document |
| `PK` (50 4B) | XLSX | ZIP format |

This allows:
- `.xls` files that are actually `.xlsx` format to be handled correctly
- `.xlsx` files that are actually `.xls` format to be handled correctly
- Automatic selection of the correct processor (xlrd vs openpyxl)

## 💡 Windows Terminal Optimization

Automatically detects and optimizes for Windows Terminal:

```python
# Auto-detects these environment variables
WT_SESSION, WT_PROFILE_ID

# Windows Terminal: Uses native UTF-8 (no configuration needed)
# Traditional cmd/PowerShell: Auto-sets chcp 65001
```

## 📄 Output Reports

After migration, three types of reports are generated:

1. **migration_report.md** - Human-readable Markdown report
   - Statistics (success/fail/success rate)
   - List of successful files
   - List of failed files (with error details)
   - Summary of main changes

2. **migration_report.json** - Machine-readable JSON report
   - Timestamps and complete statistics
   - Detailed change records for each file
   - Failure reasons and details

3. **jxls_migration.log** - Complete DEBUG log in UTF-8 encoding

## ⚠️ Known Limitations

1. **varStatus** - JXLS 2.x no longer supports varStatus (must be implemented manually in Java code)
2. **Encrypted files** - Cannot process password-protected Excel files
3. **Corrupted files** - Cannot process damaged Excel files
4. **Special formats** - Very rare special formats may require manual adjustment

## 🆘 Troubleshooting

### Q: Missing xlrd library error
```bash
pip install xlrd openpyxl
```

### Q: File has .xls extension but is actually .xlsx format
This is normal! The tool auto-detects and handles this. Use `--keep-extension` to preserve original extension.

### Q: Incorrect lastCell parameter
Manually open the .xlsx file, right-click the comment, and modify the lastCell parameter (usually the last column and last row of the table).

### Q: Some templates failed to migrate
Check the log file for specific errors. Common causes:
- File corruption
- Password protection
- Special JXLS instructions
- Complex nested structures

### Q: How to handle encrypted Excel files
The tool cannot process encrypted files. Decrypt them first.

### Q: Windows Terminal not detected
Check for environment variables `WT_SESSION` or `WT_PROFILE_ID`.

## 📝 Migration Best Practices

1. **Backup original files** - Use `_backup` suffix for backup directory
2. **Run dry-run first** - Use `--dry-run` to preview changes
3. **Test key templates** - Verify business functions after migration
4. **Review reports** - Check migration_report.md for details
5. **Validate exports** - Ensure all export functions work correctly

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Development Setup
```bash
# Clone the repository
git clone https://github.com/your-org/jxls-migration-tool.git
cd jxls-migration-tool

# Install development dependencies
pip install -r requirements-dev.txt

# Run tests
python -m pytest tests/
```

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

## 👨‍💻 Author

**fivefish**
- Version: 3.0
- Date: 2025-11-07

## 🙏 Acknowledgments

- [JXLS Project](https://jxls.sourceforge.net/) - The excellent Java library for Excel templating
- [xlrd](https://github.com/python-excel/xlrd) - Python library for reading .xls files
- [openpyxl](https://openpyxl.readthedocs.io/) - Python library for reading/writing .xlsx files

## 📚 Related Projects

- [JXLS Official Documentation](https://jxls.sourceforge.net/)
- [POI Project](https://poi.apache.org/) - Apache POI - Java API for Microsoft Documents

---

**Happy Migrating! 🎉**
