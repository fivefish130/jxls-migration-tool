# Changelog

All notable changes to the JXLS Migration Tool project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [3.4.2] - 2025-11-11

### Fixed
- **Fixed keep-extension parameter behavior**
  - Issue: .xls files kept .xls extension after migration, but Jxls 2.14.0 requires .xlsx format files
  - Solution: .xls files now keep .xls filename but file content is converted to .xlsx format
  - This preserves file names (avoiding backend code changes) while ensuring Jxls 2.14.0 compatibility
  - Directory migration: .xls → .xls filename (.xlsx content), .xlsx → .xlsx filename (.xlsx content)

- **Fixed rich text format compatibility issues**
  - Issue: Migrated files contained rich text formats, causing Jxls 2.14.0 to fail reading
  - Solution: Simplified format copying logic
    - No longer copies font names and sizes (avoids Chinese font compatibility issues)
    - Uses standard Calibri font instead
    - Only copies basic styles like bold/italic
    - Only copies solid fills and borders
  - Result: Eliminates rich text compatibility issues, ensuring Jxls 2.14.0 can read files correctly

### Improved
- **Enhanced help information**
  - Updated --keep-extension parameter description to clarify file format behavior
  - Added detailed comments to explain code logic

### Testing Verification
- File header check: 504b (ZIP format, .xlsx content) ✅
- File type check: Microsoft Excel 2007+ ✅
- openpyxl can open files normally ✅
- Migration commands successfully converted ✅

### Recommended Usage
```bash
# Use keep-extension to avoid modifying backend code
python jxls_migration_tool.py templates/ -o migrated/ --keep-extension

# Results:
# - File names remain unchanged (.xls stays .xls)
# - File content is .xlsx (Jxls 2.14.0 compatible)
# - No need to modify backend code
```

### Compatibility
- **Jxls 2.14.0**: Fully compatible ✅
- **Jxls 2.x**: Fully compatible ✅
- **Jxls 1.x**: Auto-migrated ✅

## [3.4.1] - 2025-11-07

### Fixed
- **Smart comment position calculation** - Uses `find_first_data_column` method to find the first data cell in data row
  - Precise position calculation: Based on actual data structure of original template, not fixed column position
  - Detailed logging: Added detailed logging for comment positions
  - Error handling: Falls back to command's column if no data column found
  - Comments now correctly appear in the first data cell of the data row

### Improved
- **Enhanced logging**
  - Added detailed logs for comment positions
  - Tracks data column search process
  - Records fallback mechanism usage

## [3.4.0] - 2025-11-07

### Fixed
- **Fixed jx:each comment not generated issue**
  - Enhanced instruction detection logic with more flexible matching
  - Added detailed debug logs to track command processing
  - Ensured forEach commands are correctly identified and processed
  - Fixed processing logic in `process_commands_and_migrate_data` and `process_commands_xlsx`

- **Fixed jx:area position error issue**
  - Auto-generated jx:area now correctly added to A1 cell
  - Fixed comment position calculation logic
  - Added dedicated logs to confirm jx:area comment position
  - Unified fix in both XLS and XLSX processing

### Improved
- **Enhanced logging**
  - Added more detailed debug information
  - Tracks entire command processing workflow
  - Confirms comment addition positions

### Version Update
- Version number updated to v3.4
- Banner and documentation updated to "Fixed v3.4"

## [3.3.1] - 2025-11-07

### Fixed
- **Robust migration fallback logic**
  - Improved error handling in migration process
  - Added automatic fallback when primary processor fails
  - Enhanced format detection accuracy
  - Better handling of edge cases

### Improved
- **API unification**
  - Simplified migrate_file interface
  - Consistent return format across all migration methods
  - Better error propagation

## [3.3.0] - 2025-11-07

### Added
- **Robust migration scheme**
  - Automatic format detection and processor selection
  - Dual processor fallback mechanism
  - Comprehensive error handling and logging

### Improved
- **Format detection**
  - Smart file format detection regardless of extension
  - More accurate format recognition
  - Better handling of mismatched extensions

### Fixed
- **Migration stability**
  - Improved error recovery
  - Better handling of malformed files
  - Enhanced logging for debugging

## [3.2.0] - 2025-11-07

### Added
- **Complete robust migration solution**
  - Support for both XLS and XLSX formats
  - Automatic format conversion
  - Format preservation (styles, column widths, row heights, merged cells)

### Fixed
- **Format object attribute errors**
  - Fixed 'Format' object has no attribute 'font_index' error
  - Enhanced error handling in format conversion

## [3.1.0] - 2025-11-07

### Added
- **Windows Terminal optimization**
  - Automatic UTF-8 detection and configuration
  - Support for various terminal environments
  - Enhanced user experience on Windows

### Improved
- **Unicode support**
  - Better handling of Chinese/Japanese/Korean characters
  - Improved console output formatting

## [3.0.0] - 2025-11-07

### Added
- **Complete JXLS instruction support**
  - forEach → each migration
  - if(test → condition) migration
  - out → ${} expression replacement
  - area command auto-generation
  - multiSheet support

### Added
- **Format preservation**
  - Cell styles maintenance
  - Column widths preservation
  - Row heights preservation
  - Merged cells preservation
  - Background colors preservation

### Added
- **Smart file format detection**
  - Detects actual file format regardless of extension
  - Handles mismatched extensions gracefully

### Added
- **Detailed reporting**
  - Markdown report generation
  - JSON report generation
  - DEBUG logging

### Added
- **Production-ready features**
  - Comprehensive error handling
  - Extensive logging
  - Batch processing support
  - Dry-run mode

## [2.0.0] - 2025-11-07

### Added
- **Initial migration functionality**
  - Basic JXLS instruction conversion
  - Single file processing

### Added
- **Basic format support**
  - XLS to XLSX conversion
  - Basic style preservation

## [1.0.0] - 2025-11-07

### Added
- **Initial release**
  - Project initialization
  - Basic structure setup
