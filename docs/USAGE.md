# JXLS Migration Tool - Usage Guide

This guide provides detailed instructions on how to use the JXLS Migration Tool effectively.

## Table of Contents

- [Installation](#installation)
- [Basic Usage](#basic-usage)
- [Command Line Options](#command-line-options)
- [Migration Modes](#migration-modes)
- [File Format Handling](#file-format-handling)
- [Report Generation](#report-generation)
- [Error Handling](#error-handling)
- [Advanced Scenarios](#advanced-scenarios)
- [Best Practices](#best-practices)
- [FAQ](#faq)

## Installation

### System Requirements

- **Python**: 3.6 or higher (Python 3.12 recommended)
- **Operating System**: Windows 10/11, Linux, macOS
- **Memory**: Minimum 512MB RAM (1GB+ recommended for large files)
- **Disk Space**: 2x the size of your Excel files (for processing)

### Install Dependencies

#### Using pip
```bash
pip install xlrd==2.0.1 openpyxl
```

#### Using conda
```bash
conda install xlrd=2.0.1 openpyxl
```

#### Verify Installation
```bash
python -c "import xlrd, openpyxl; print('Dependencies installed successfully')"
```

### Python Environment Setup

#### Using virtualenv
```bash
# Create virtual environment
python -m venv jxls-migration-env

# Activate on Windows
jxls-migration-env\Scripts\activate

# Activate on Linux/macOS
source jxls-migration-env/bin/activate

# Install dependencies
pip install xlrd==2.0.1 openpyxl
```

#### Using conda
```bash
# Create conda environment
conda create -n jxls-migration python=3.12
conda activate jxls-migration
pip install xlrd==2.0.1 openpyxl
```

## Basic Usage

### Single File Migration

#### Convert a single file
```bash
# Convert input.xls to output.xlsx
python jxls_migration_tool.py input.xls -f output.xlsx

# Convert with keep-extension
python jxls_migration_tool.py input.xls -f output.xls --keep-extension
```

#### What happens:
1. Reads the input file
2. Detects JXLS instructions
3. Converts them to JXLS 2.x format
4. Preserves all formatting
5. Writes the output file

### Directory Migration

#### Migrate all Excel files in a directory
```bash
# Basic directory migration
python jxls_migration_tool.py /path/to/excel/templates

# Migrate to specific output directory
python jxls_migration_tool.py /path/to/excel/templates -o /path/to/output

# Keep original file extensions
python jxls_migration_tool.py /path/to/excel/templates -o /path/to/output --keep-extension

# With verbose logging
python jxls_migration_tool.py /path/to/excel/templates --keep-extension --verbose
```

#### What happens:
1. Scans the directory for Excel files (.xls, .xlsx)
2. Detects JXLS instructions in each file
3. Processes files with JXLS instructions
4. Skips files without JXLS instructions
5. Generates migration reports
6. Creates logs

### Dry Run Mode

#### Preview changes without modifying files
```bash
# Dry run with verbose output
python jxls_migration_tool.py /path/to/excel/templates --dry-run --verbose

# Dry run with output directory specified
python jxls_migration_tool.py /path/to/excel/templates -o /path/to/output --dry-run
```

**What you get:**
- List of files that would be processed
- Number of JXLS commands found in each file
- Estimated conversion details
- No actual file modifications

## Command Line Options

### Full Option Reference

```bash
python jxls_migration_tool.py [OPTIONS] INPUT_PATH
```

#### Positional Arguments

- `INPUT_PATH` (required)
  - Type: String
  - Description: Path to input file or directory
  - Examples:
    - `input.xls` (single file)
    - `input_directory` (directory containing Excel files)

#### Optional Arguments

##### `-h, --help`
- Description: Show help message and exit
- Example: `python jxls_migration_tool.py --help`

##### `-o OUTPUT, --output OUTPUT`
- Description: Output directory for migrated files
- Default: Same as input directory
- Type: String
- Example: `python jxls_migration_tool.py input_dir -o output_dir`

##### `-f OUTPUT_FILE, --file OUTPUT_FILE`
- Description: Output file (for single file migration only)
- Default: Not applicable
- Type: String
- Example: `python jxls_migration_tool.py input.xls -f output.xlsx`

##### `--keep-extension`
- Description: Keep original file extensions (.xls → .xls, .xlsx → .xlsx)
- Default: Convert all .xls to .xlsx
- Type: Boolean flag (no value needed)
- Example: `python jxls_migration_tool.py input_dir --keep-extension`

##### `--dry-run`
- Description: Preview changes without modifying files
- Default: False
- Type: Boolean flag
- Example: `python jxls_migration_tool.py input_dir --dry-run`

##### `-v, --verbose`
- Description: Enable verbose logging
- Default: False
- Type: Boolean flag
- Example: `python jxls_migration_tool.py input_dir --verbose`

### Option Combinations

#### Common Use Cases

```bash
# 1. Standard migration with extensions preserved
python jxls_migration_tool.py input_dir -o output_dir --keep-extension

# 2. Preview before migration
python jxls_migration_tool.py input_dir --dry-run --verbose

# 3. In-place migration (overwrite original files)
python jxls_migration_tool.py input_dir --keep-extension

# 4. Convert to .xlsx (no extension preservation)
python jxls_migration_tool.py input_dir -o output_dir

# 5. Single file with specific output name
python jxls_migration_tool.py input.xls -f specific_output.xlsx
```

## Migration Modes

### Mode 1: Standard Conversion (.xls → .xlsx)

**Purpose**: Convert all .xls files to .xlsx format

**Command**:
```bash
python jxls_migration_tool.py input_dir -o output_dir
```

**Behavior**:
- All .xls files → .xlsx
- All .xlsx files remain .xlsx
- JXLS instructions converted
- Formatting preserved

**When to use**:
- Standard migration
- Want all files in modern .xlsx format
- Don't care about keeping original extensions

### Mode 2: Extension Preservation

**Purpose**: Keep original file extensions

**Command**:
```bash
python jxls_migration_tool.py input_dir --keep-extension
```

**Behavior**:
- .xls files remain .xls
- .xlsx files remain .xlsx
- JXLS instructions converted
- Formatting preserved
- Smart format detection (handles mismatched extensions)

**When to use**:
- When file extension is part of your workflow
- When you need to preserve original file types
- When you want to minimize changes

### Mode 3: In-Place Migration

**Purpose**: Migrate files in the same directory

**Command**:
```bash
python jxls_migration_tool.py input_dir --keep-extension
```

**Behavior**:
- Files overwritten in place
- Original files replaced
- Backup recommended before using this mode

**When to use**:
- When you have backups
- When you're comfortable with overwriting files
- Testing environment

### Mode 4: Preview Mode

**Purpose**: See what would be migrated without making changes

**Command**:
```bash
python jxls_migration_tool.py input_dir --dry-run --verbose
```

**Behavior**:
- Scans files
- Reports on JXLS instructions found
- Shows what would be converted
- No file modifications

**When to use**:
- First time using the tool
- Audit before migration
- Troubleshooting
- Learning how the tool works

## File Format Handling

### Smart Format Detection

The tool uses file header analysis to detect the actual format, not just the extension:

#### Detection Logic

```python
# Read first 8 bytes of file
with open(filepath, 'rb') as f:
    header = f.read(8)

# Check header
if header.startswith(b'\xd0\xcf\x11\xe0'):
    # OLE2 format = .xls
    actual_format = 'xls'
elif header.startswith(b'PK'):
    # ZIP format = .xlsx
    actual_format = 'xlsx'
```

#### Format Mismatches Handled

**Scenario 1**: File has .xls extension but is actually .xlsx
- Detected correctly
- Processed as .xlsx
- Output respects `--keep-extension` flag

**Scenario 2**: File has .xlsx extension but is actually .xls
- Detected correctly
- Processed as .xls
- Output respects `--keep-extension` flag

**Log Output**:
```
⚠️  File has .xls extension but actual format is .xlsx
   Detected format: xlsx
   Processing as: xlsx
```

### Supported Formats

| Format | Extension | Detection | Processor |
|--------|-----------|-----------|-----------|
| Excel 97-2003 | .xls | OLE2 header (D0CF11E0) | xlrd |
| Excel 2007+ | .xlsx | ZIP header (PK) | openpyxl |

### Unsupported Formats

- **.xlsb** (Excel Binary Workbook)
- **.xlam** (Excel Add-in)
- **.xlsm** (Excel Macro-enabled Workbook) - *can read but macros not preserved*
- **Password-protected files** - *cannot read*
- **Corrupted files** - *will fail*

## Report Generation

### Automatic Report Generation

After migration, three types of reports are generated:

#### 1. Markdown Report (migration_report.md)

**Location**: Output directory
**Purpose**: Human-readable summary
**Content**:
- Header with version and date
- Statistics (total, success, failed, success rate)
- List of processed files
- List of successful conversions
- List of failed conversions (with error details)
- Summary of main changes
- Next steps recommendations

**Example**:
```markdown
# JXLS Migration Report

## Statistics
- Total files: 50
- Successfully migrated: 50
- Failed: 0
- Success rate: 100%

## Successfully Migrated Files
- purchase_order_export.xls
  - JXLS commands found: 2
  - Commands converted: 2
  - Sheet: 采购单

## Failed Files
(None)
```

#### 2. JSON Report (migration_report.json)

**Location**: Output directory
**Purpose**: Machine-readable detailed data
**Content**:
- Timestamp
- Complete statistics
- Detailed records for each file
- Before/after file information
- Conversion details
- Error information (if any)

**Example**:
```json
{
  "timestamp": "2025-11-07 10:00:00",
  "statistics": {
    "total_files": 50,
    "successful": 50,
    "failed": 0,
    "success_rate": 100.0
  },
  "files": [
    {
      "filename": "purchase_order_export.xls",
      "status": "success",
      "commands_found": 2,
      "commands_converted": 2,
      "sheets": ["采购单"]
    }
  ]
}
```

#### 3. Debug Log (jxls_migration.log)

**Location**: Output directory
**Purpose**: Complete execution log
**Content**:
- Timestamped log entries
- DEBUG level details
- Error stack traces
- File processing details
- Command detection and conversion logs

**Example**:
```
2025-11-07 10:00:00 - INFO - Starting migration of directory: input_dir
2025-11-07 10:00:01 - DEBUG - Processing file: purchase_order_export.xls
2025-11-07 10:00:01 - DEBUG - Detected format: xlsx
2025-11-07 10:00:01 - DEBUG - Found 2 JXLS commands
2025-11-07 10:00:02 - INFO - Successfully migrated: purchase_order_export.xls
```

### Custom Report Location

Reports are always generated in the output directory. To specify a different location:

```bash
# Reports will be in output_dir
python jxls_migration_tool.py input_dir -o output_dir --keep-extension
```

## Error Handling

### Common Errors and Solutions

#### Error: "xlrd library not found"

**Solution**:
```bash
pip install xlrd==2.0.1
```

#### Error: "Permission denied"

**Cause**: No write permission to output directory

**Solution**:
- Check directory permissions
- Run with appropriate user permissions
- Specify a different output directory

#### Error: "File is encrypted"

**Cause**: Excel file is password-protected

**Solution**:
- Decrypt the file manually
- Remove password protection
- Re-run migration

#### Error: "File is corrupted"

**Cause**: Excel file is damaged or invalid

**Solution**:
- Verify file can be opened in Excel
- Restore from backup
- Re-create file if necessary

#### Error: "lastCell parameter incorrect"

**Cause**: Complex nested structures may result in incorrect lastCell

**Solution**:
- Manual review of migrated file
- Adjust lastCell in Excel comments
- Report as issue for tool improvement

### Error Recovery

The tool continues processing even if individual files fail:

```bash
# Example output with failures
✅ Successfully migrated: 48 files
❌ Failed: 2 files
📊 Total: 50 files
🎯 Success rate: 96.00%
```

Check the reports to see which files failed and why.

## Advanced Scenarios

### Large File Migration

For directories with many large files:

```bash
# Process with verbose logging
python jxls_migration_tool.py large_directory --keep-extension --verbose

# Dry run first to estimate time
python jxls_migration_tool.py large_directory --dry-run
```

**Performance Tips**:
- Ensure sufficient disk space (2x file size)
- Use SSD storage for better performance
- Close other applications to free memory

### Mixed File Types

Directory with various Excel files:

```bash
python jxls_migration_tool.py mixed_directory --keep-extension --verbose
```

The tool will:
- Process .xls/.xlsx files with JXLS instructions
- Skip files without JXLS instructions
- Skip non-Excel files
- Generate comprehensive reports

### Automation Script Integration

```bash
#!/bin/bash
# migrate.sh - Batch migration script

INPUT_DIR=$1
OUTPUT_DIR=$2

if [ -z "$INPUT_DIR" ] || [ -z "$OUTPUT_DIR" ]; then
    echo "Usage: $0 <input_dir> <output_dir>"
    exit 1
fi

# Create output directory
mkdir -p "$OUTPUT_DIR"

# Run migration
python jxls_migration_tool.py "$INPUT_DIR" -o "$OUTPUT_DIR" --keep-extension --verbose

# Check results
if [ $? -eq 0 ]; then
    echo "Migration completed successfully"
    echo "Reports available in: $OUTPUT_DIR"
else
    echo "Migration failed. Check logs."
    exit 1
fi
```

### Git Integration

```bash
# Add to .gitignore
migration_report.*
jxls_migration.log

# Pre-commit hook example
#!/bin/bash
# Run dry-run on changed files
git diff --name-only | grep '\.xls$|\.xlsx$' | while read file; do
    echo "Checking $file for JXLS 1.x syntax..."
    if grep -q '<jx:' "$file"; then
        echo "ERROR: $file contains JXLS 1.x syntax. Run migration tool."
        exit 1
    fi
done
```

## Best Practices

### 1. Always Backup First

```bash
# Create backup
cp -r templates templates_backup_$(date +%Y%m%d)

# Run migration
python jxls_migration_tool.py templates --keep-extension
```

### 2. Use Dry-Run First

```bash
# Always preview changes
python jxls_migration_tool.py templates --dry-run --verbose

# Review output
cat migration_report.md
```

### 3. Test Key Templates

After migration, test critical templates to ensure they work:
- Run your application
- Test export functions
- Verify output matches expectations
- Check for any runtime errors

### 4. Review Reports

Always check generated reports:
- `migration_report.md` - Human summary
- `migration_report.json` - Detailed data
- `jxls_migration.log` - Debug information

### 5. Version Control

```bash
# Commit migrated files
git add -A
git commit -m "Migrate JXLS templates to 2.14.0"

# Tag the version
git tag -a v2.14.0 -m "JXLS 2.14.0 migration"
```

### 6. Document Your Migration

```markdown
# Migration Record

Date: 2025-11-07
Version: 3.0
Files Migrated: 50
Success Rate: 100%

Files Tested:
- [ ] purchase_order_export.xls
- [ ] inventory_report.xlsx
...

Issues Found:
- (List any issues and their solutions)
```

### 7. Gradual Rollout

```bash
# Phase 1: Test environment
python jxls_migration_tool.py test_templates -o test_output --keep-extension

# Phase 2: Staging environment
python jxls_migration_tool.py staging_templates -o staging_output --keep-extension

# Phase 3: Production
python jxls_migration_tool.py prod_templates -o prod_output --keep-extension
```

## FAQ

### Q: How long does migration take?

**A**: Typically < 1 second per file. For 50 files, expect < 30 seconds. Large files may take longer.

### Q: Will my formulas be preserved?

**A**: Yes, cell formulas (${...} expressions) are preserved. The tool only modifies JXLS instructions.

### Q: What about images in my templates?

**A**: Images are preserved. The tool maintains all non-text content.

### Q: Can I migrate back from 2.x to 1.x?

**A**: No, the migration is one-way. Keep backups of your original files.

### Q: Does the tool work with Chinese/Japanese/Korean characters?

**A**: Yes, full Unicode support. Windows Terminal is auto-detected and configured.

### Q: What if I have a custom JXLS instruction?

**A**: The tool supports standard JXLS instructions (forEach, if, out, area, multiSheet). Custom instructions need manual migration.

### Q: How do I verify migration success?

**A**:
1. Check the migration report
2. Open migrated files in Excel
3. Verify JXLS comments are present
4. Test template rendering
5. Run your application tests

### Q: Can I use this in a CI/CD pipeline?

**A**: Yes! Example for GitHub Actions:

```yaml
- name: Migrate JXLS templates
  run: |
    python jxls_migration_tool.py templates --keep-extension
    # Check for failures
    if grep -q "Failed:" migration_report.md; then
      echo "Migration failed!"
      exit 1
    fi
```

### Q: Where are the reports?

**A**: Reports are in the output directory:
- `migration_report.md` (Markdown summary)
- `migration_report.json` (Detailed JSON)
- `jxls_migration.log` (Debug log)

### Q: How do I report an issue?

**A**:
1. Enable `--verbose` flag
2. Include the log file (`jxls_migration.log`)
3. Include the migration report
4. Provide sample file if possible
5. Report tool version and Python version

---

**Need more help?** Check out [API Documentation](API.md) or visit the [project repository](https://github.com/your-org/jxls-migration-tool).
