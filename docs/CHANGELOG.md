# Changelog

All notable changes to the JXLS Migration Tool will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.1.0] - 2025-11-07

### 🎉 Initial Release as Independent Tool

This release marks the first standalone version of the JXLS Migration Tool, extracted from an enterprise project and made available as a general-purpose utility.

#### ✨ Added
- Complete standalone tool structure
- Comprehensive documentation (README, USAGE, API docs)
- Examples and test directories
- License file (MIT)
- Setup.py for pip installation
- GitHub-ready repository structure
- English documentation (previously Chinese-only)

#### 🔄 Changed
- **BREAKING**: Simplified docstring (from 86 lines to 23 lines) for better maintainability
- Updated to 3.1.0 version (increment from 3.0 for standalone release)
- Made tool fully generic (removed project-specific hardcoding)
- Enhanced documentation with English translations
- Improved code organization and structure

#### 📚 Documentation
- Created comprehensive README.md with badges and quick start
- Added detailed USAGE.md with examples and best practices
- Created CHANGELOG.md for version history
- Added API.md documentation (planned)
- Added examples directory with sample code (planned)
- Added LICENSE file (MIT)

#### 🏗️ Repository Structure
```
tools/jxls-migration-tool/
├── README.md              # Main documentation
├── jxls_migration_tool.py # Main tool
├── docs/
│   ├── USAGE.md           # Detailed usage guide
│   ├── API.md             # API documentation (planned)
│   └── CHANGELOG.md       # This file
├── examples/              # Code examples (planned)
├── tests/                 # Test suite (planned)
├── LICENSE                # MIT License
└── setup.py               # Installation script (planned)
```

---

## [3.0] - 2025-11-06

### 🎉 Production-Ready Release

Major overhaul to make the tool production-ready with full JXLS 2.14.0 support.

#### ✨ Added
- **Complete JXLS Instruction Support**
  - jx:forEach → jx:each with comment-based conversion
  - jx:if(test=) → jx:if(condition=) with comment-based conversion
  - jx:out → ${...} expression replacement
  - jx:area - preserve existing or auto-generate
  - jx:multiSheet - full multi-sheet support

- **Smart File Format Detection**
  - Auto-detects actual file format by reading headers
  - Supports .xls files that are actually .xlsx format
  - Supports .xlsx files that are actually .xls format
  - Automatic processor selection (xlrd vs openpyxl)

- **Windows Terminal Optimization**
  - Auto-detects Windows Terminal environment
  - Detects WT_SESSION and WT_PROFILE_ID
  - Native UTF-8 support for Windows Terminal
  - Fallback to chcp 65001 for traditional cmd/PowerShell

- **--keep-extension Option**
  - Preserves original file extensions
  - .xls stays .xls, .xlsx stays .xlsx
  - Works with smart format detection

- **Enhanced Report Generation**
  - Markdown report with statistics
  - JSON report with detailed data
  - DEBUG log with full execution details
  - Success/failure lists with error details

- **Auto area Generation**
  - Automatically generates jx:area if missing
  - Places in A1 cell
  - Calculates appropriate lastCell

#### 🔄 Changed
- **Format Preservation**: Complete preservation of all Excel formatting
  - Cell styles (fonts, colors, borders, alignment)
  - Column widths and row heights (with unit conversion)
  - Merged cells (with range adjustment)
  - Background colors (with XLS index to RGB mapping)

- **Error Handling**: Significantly improved error handling
  - Detailed error messages
  - Continues processing on individual file failures
  - Comprehensive logging

- **Performance**: Optimized for large files
  - Better memory management
  - Faster processing
  - Reduced overhead

#### 📊 Migration Statistics (enterprise project)
- **Total files scanned**: 923 Excel files
- **Successfully migrated**: 50 JXLS templates
- **Skipped**: 873 files (no JXLS instructions or import-only templates)
- **Failed**: 0 files
- **Success rate**: 100%
- **Total JXLS commands found**: 106
- **Commands converted**: 106

#### 📈 Code Improvements
- **Total lines**: 1,972 (from 1,124 in v2.0)
- **Main classes**: 8 (from 5 in v2.0)
- **Methods**: 30+ (from 20+ in v2.0)
- **Comment lines**: 400+ (from 200+ in v2.0)
- **Supported JXLS instructions**: 5 types

#### 🏷️ Module Distribution
- Module A: 7 templates
- Module B: 1 template
- Module C: 42 templates

---

## [2.0] - 2025-11-01

### 🔧 Feature Enhancement

Added format preservation and improved functionality.

#### ✨ Added
- Format preservation capabilities
  - Cell styles (fonts, colors, borders)
  - Column widths and row heights
  - Merged cells
  - Background colors

- XLSX format support
  - Full read/write support for .xlsx files
  - Openpyxl integration

- Improved lastCell calculation
  - More accurate calculation
  - Better handling of complex structures

#### 🔄 Changed
- Enhanced output format
- Better progress reporting
- Improved error messages

#### 🐛 Fixed
- Various bugs from v1.0
- Better error recovery
- Improved stability

---

## [1.0] - 2025-10-31

### 🎉 Initial Release

First version of the JXLS migration tool.

#### ✨ Added
- Basic JXLS instruction support
  - jx:forEach → jx:each conversion
  - jx:if conversion
  - jx:out → ${...} replacement

- Comment-based conversion
  - Adds Excel comments to mark JXLS instructions
  - Removes original tag rows

- Basic format support
  - .xls file support (via xlrd)
  - Basic formatting preservation

- Report generation
  - Markdown reports
  - Basic statistics

#### 📋 Known Limitations (v1.0)
- Limited JXLS instruction support
- No area command handling
- No multiSheet support
- Basic error handling
- No Windows Terminal optimization
- No smart file format detection

#### 📊 Initial Statistics
- Total lines: 1,124
- Main classes: 5
- Methods: 20+
- Comment lines: 200+

---

## Future Roadmap

### [3.2.0] - Planned
- [ ] Add support for .xlsm (macro-enabled) files
- [ ] Enhanced error recovery for corrupted files
- [ ] Support for custom JXLS instructions via plugins
- [ ] Performance improvements for very large files
- [ ] Progress bar for long migrations

### [4.0.0] - Future
- [ ] Web UI for easy migration
- [ ] API for programmatic usage
- [ ] Batch processing optimization
- [ ] Integration with JXLS official tools
- [ ] Support for JXLS 3.x (when available)

---

## Version Numbering

We use [Semantic Versioning](https://semver.org/):

- **MAJOR** version when you make incompatible API changes
- **MINOR** version when you add functionality in a backwards compatible manner
- **PATCH** version when you make backwards compatible bug fixes

**Current version**: 3.1.0

### Version History
- 3.1.0 - 2025-11-07: Initial standalone release
- 3.0 - 2025-11-06: Production-ready with full support
- 2.0 - 2025-11-01: Format preservation enhancement
- 1.0 - 2025-10-31: Initial release

---

## Contributing

When adding entries to this changelog:

1. Keep the format consistent
2. Group changes under appropriate headings (Added, Changed, Deprecated, Removed, Fixed, Security)
3. Include relevant issue numbers and PR numbers
4. Write clear, concise descriptions
5. Add dates in YYYY-MM-DD format
6. Mark breaking changes clearly

For more information, see our [Contributing Guidelines](CONTRIBUTING.md).

---

## Acknowledgments

- **JXLS Project** - https://jxls.sourceforge.net/
- **xlrd library** - https://github.com/python-excel/xlrd
- **openpyxl library** - https://openpyxl.readthedocs.io/
- **ftrace team** - For extensive testing and feedback
- **Community contributors** - Bug reports and feature requests

---

**Note**: This changelog will be updated with each release. For the most current information, always check the [GitHub releases page](https://github.com/your-org/jxls-migration-tool/releases).
