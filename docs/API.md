# JXLS Migration Tool - API Documentation

This document provides detailed information about the programmatic API of the JXLS Migration Tool.

## Table of Contents

- [Overview](#overview)
- [Main Class: JxlsMigrationTool](#main-class-jxlsmigrationtool)
- [Command Classes](#command-classes)
- [Helper Classes](#helper-classes)
- [Utility Functions](#utility-functions)
- [Error Handling](#error-handling)
- [Examples](#examples)
- [Type Hints](#type-hints)

## Overview

The JXLS Migration Tool provides a Python API for programmatic migration of JXLS 1.x templates to JXLS 2.14.0. The API is designed to be:

- **Simple**: Easy to use for basic migrations
- **Flexible**: Configurable for advanced use cases
- **Robust**: Comprehensive error handling and logging
- **Extensible**: Custom command processors can be added

## Main Class: JxlsMigrationTool

The `JxlsMigrationTool` class is the primary interface for migration operations.

### Constructor

```python
class JxlsMigrationTool:
    def __init__(self, verbose=False, log_file=None):
        """
        Initialize the migration tool.

        Args:
            verbose (bool): Enable verbose logging. Default: False.
            log_file (str, optional): Path to log file. If None, logs to console.
        """
```

**Example**:
```python
from jxls_migration_tool import JxlsMigrationTool

# Basic initialization
tool = JxlsMigrationTool()

# With verbose logging
tool = JxlsMigrationTool(verbose=True)

# With custom log file
tool = JxlsMigrationTool(verbose=True, log_file="migration.log")
```

### Method: migrate_directory()

Migrate all Excel files in a directory.

```python
def migrate_directory(
    self,
    input_dir,
    output_dir,
    keep_extension=True,
    dry_run=False
) -> bool:
    """
    Migrate all Excel files in a directory.

    Args:
        input_dir (str): Path to input directory.
        output_dir (str): Path to output directory.
        keep_extension (bool): Keep original file extensions. Default: True.
        dry_run (bool): Preview changes without modifying files. Default: False.

    Returns:
        bool: True if migration completed (may have partial failures),
              False if critical error occurred.

    Raises:
        FileNotFoundError: If input_dir does not exist.
        PermissionError: If no write permission to output_dir.
        Exception: For other critical errors.
    """
```

**Example**:
```python
tool = JxlsMigrationTool(verbose=True)

success = tool.migrate_directory(
    input_dir="/path/to/templates",
    output_dir="/path/to/output",
    keep_extension=True,
    dry_run=False
)

if success:
    print("Migration completed successfully")
else:
    print("Migration encountered errors")
```

**Return Value**:
- `True`: Migration completed (check reports for individual file status)
- `False`: Critical error prevented migration

### Method: migrate_file()

Migrate a single Excel file.

```python
def migrate_file(
    self,
    input_path,
    output_path=None,
    keep_extension=True
) -> bool:
    """
    Migrate a single Excel file.

    Args:
        input_path (str): Path to input Excel file.
        output_path (str, optional): Path to output file.
            If None, uses input_path. Default: None.
        keep_extension (bool): Keep original file extension. Default: True.

    Returns:
        bool: True if migration successful, False otherwise.

    Raises:
        FileNotFoundError: If input_path does not exist.
        ValueError: If file is not a valid Excel file.
        Exception: For other errors during migration.
    """
```

**Example**:
```python
tool = JxlsMigrationTool()

# Migrate to specific output file
success = tool.migrate_file(
    input_path="template.xls",
    output_path="migrated_template.xls",
    keep_extension=True
)

# Migrate in place
success = tool.migrate_file(
    input_path="template.xls",
    keep_extension=True
)

# Migrate with extension change
success = tool.migrate_file(
    input_path="template.xls",
    output_path="template.xlsx",
    keep_extension=False
)
```

**Return Value**:
- `True`: File migrated successfully
- `False`: Migration failed (check logs for details)

### Method: generate_report()

Generate migration reports in Markdown and JSON formats.

```python
def generate_report(self, output_dir) -> None:
    """
    Generate migration reports.

    Args:
        output_dir (str): Directory where reports will be saved.
    """
```

**Example**:
```python
tool = JxlsMigrationTool(verbose=True)

# Run migration
tool.migrate_directory("input", "output")

# Generate reports
tool.generate_report("output")
```

**Output Files**:
- `output/migration_report.md` - Human-readable summary
- `output/migration_report.json` - Machine-readable data
- `output/jxls_migration.log` - DEBUG log

## Command Classes

Command classes handle parsing and conversion of specific JXLS instructions.

### JxlsCommand (Base Class)

```python
class JxlsCommand:
    def __init__(self, content, row):
        """
        Initialize base JXLS command.

        Args:
            content (str): Cell content containing JXLS instruction.
            row (int): Row number where command was found.
        """
        self.content = content
        self.row = row
        self.instruction = None

    def parse(self) -> dict:
        """
        Parse JXLS instruction.

        Returns:
            dict: Parsed instruction data.
        """
        raise NotImplementedError("Subclasses must implement parse()")

    def convert(self) -> str:
        """
        Convert to JXLS 2.x format.

        Returns:
            str: Converted instruction.
        """
        raise NotImplementedError("Subclasses must implement convert()")
```

### ForEachCommand

Handles `jx:forEach` instruction conversion.

```python
class ForEachCommand(JxlsCommand):
    def parse(self) -> dict:
        """
        Parse forEach instruction.

        Returns:
            dict: Parsed with keys: items, var, lastCell, etc.
        """
        # Implementation
        pass

    def convert(self) -> str:
        """
        Convert to jx:each with comment format.

        Returns:
            str: JXLS 2.x instruction.
        """
        # Example output: jx:each(items="datas" var="item" lastCell="C3")
```

**Example**:
```python
from jxls_migration_tool import ForEachCommand

# Create command from cell content
command = ForEachCommand('<jx:forEach items="datas" var="item">', row=3)

# Parse instruction
parsed = command.parse()
# Returns: {'items': 'datas', 'var': 'item', 'type': 'forEach'}

# Convert to JXLS 2.x
converted = command.convert()
# Returns: jx:each(items="datas" var="item" lastCell="C3")
```

### IfCommand

Handles `jx:if` instruction conversion.

```python
class IfCommand(JxlsCommand):
    def parse(self) -> dict:
        """
        Parse if instruction.

        Returns:
            dict: Parsed with keys: condition, lastCell, etc.
        """
        pass

    def convert(self) -> str:
        """
        Convert to jx:if with condition parameter.

        Returns:
            str: JXLS 2.x instruction.
        """
```

### OutCommand

Handles `jx:out` instruction conversion.

```python
class OutCommand(JxlsCommand):
    def parse(self) -> dict:
        """
        Parse out instruction.

        Returns:
            dict: Parsed with key: select
        """
        pass

    def convert(self) -> str:
        """
        Convert to ${...} expression.

        Returns:
            str: Converted expression.
        """
```

**Example**:
```python
# Input: <jx:out select="item.name"/>
# Output: ${item.name}
```

### AreaCommand

Handles `jx:area` instruction (v3.0+).

```python
class AreaCommand(JxlsCommand):
    def parse(self) -> dict:
        """
        Parse area instruction.

        Returns:
            dict: Parsed with key: lastCell
        """
        pass

    def convert(self) -> str:
        """
        Convert to jx:area comment format.

        Returns:
            str: JXLS 2.x instruction.
        """
```

### MultiSheetCommand

Handles `jx:multiSheet` instruction (v3.0+).

```python
class MultiSheetCommand(JxlsCommand):
    def parse(self) -> dict:
        """
        Parse multiSheet instruction.

        Returns:
            dict: Parsed with key: data
        """
        pass

    def convert(self) -> str:
        """
        Convert to jx:multiSheet comment format.

        Returns:
            str: JXLS 2.x instruction.
        """
```

## Helper Classes

### ExcelFormatConverter

Converts formatting from XLS to XLSX.

```python
class ExcelFormatConverter:
    def __init__(self):
        """Initialize format converter."""
        self.xls_color_map = {
            1: "000000",  # Black
            2: "FFFFFF",  # White
            # ... more colors
        }

    def convert_font(self, font) -> Font:
        """
        Convert XLS font to XLSX font.

        Args:
            font: xlrd font object

        Returns:
            openpyxl Font object
        """
        pass

    def convert_fill(self, fill) -> PatternFill:
        """
        Convert XLS fill to XLSX fill.

        Args:
            fill: xlrd fill object

        Returns:
            openpyxl PatternFill object
        """
        pass

    def convert_border(self, border) -> Border:
        """
        Convert XLS border to XLSX border.

        Args:
            border: xlrd border object

        Returns:
            openpyxl Border object
        """
        pass

    def get_rgb_from_xls_color(self, color_index) -> str:
        """
        Get RGB value from XLS color index.

        Args:
            color_index (int): XLS color index

        Returns:
            str: RGB hex value (e.g., "FF0000")
        """
        return self.xls_color_map.get(color_index, "000000")
```

**Example**:
```python
from jxls_migration_tool import ExcelFormatConverter

converter = ExcelFormatConverter()

# Convert a font
xlsx_font = converter.convert_font(xls_font)

# Get RGB from color index
rgb = converter.get_rgb_from_xls_color(3)  # Returns blue
```

## Utility Functions

### detect_excel_format()

Detect actual Excel file format by reading file header.

```python
def detect_excel_format(filepath: str) -> str:
    """
    Detect Excel file format by reading file header.

    Args:
        filepath (str): Path to Excel file

    Returns:
        str: 'xls' for OLE2 format, 'xlsx' for ZIP format

    Raises:
        FileNotFoundError: If file does not exist
        IOError: If file cannot be read
    """
```

**Example**:
```python
from jxls_migration_tool import detect_excel_format

# Detect format
format_type = detect_excel_format("template.xls")
print(format_type)  # Output: 'xls' or 'xlsx'
```

**Return Values**:
- `'xls'`: OLE2/Compound Document format (Excel 97-2003)
- `'xlsx'`: ZIP format (Excel 2007+)

### setup_logging()

Configure logging for the migration tool.

```python
def setup_logging(verbose=False, log_file=None) -> logging.Logger:
    """
    Set up logging configuration.

    Args:
        verbose (bool): Enable DEBUG level logging
        log_file (str, optional): Path to log file

    Returns:
        logging.Logger: Configured logger instance
    """
```

**Example**:
```python
from jxls_migration_tool import setup_logging

# Set up logging to file
logger = setup_logging(verbose=True, log_file="migration.log")

# Use logger
logger.info("Migration started")
logger.debug("Processing file: template.xls")
```

## Error Handling

The tool uses a hierarchical error handling approach:

### Error Types

#### FileNotFoundError
**Cause**: Input file or directory does not exist.
**Handling**: Tool will print error and continue with other files.

#### PermissionError
**Cause**: No read/write permissions.
**Handling**: Tool will print error and continue.

#### ValueError
**Cause**: Invalid file format or corrupt data.
**Handling**: Tool will log error and skip file.

#### Exception
**Cause**: Unexpected errors.
**Handling**: Tool will log full traceback and continue.

### Error Recovery

The tool continues processing even if individual files fail:

```python
try:
    success = tool.migrate_file("template.xls", "output.xls")
    if not success:
        print("File migration failed - check logs")
except Exception as e:
    print(f"Unexpected error: {e}")
```

### Best Practices

1. **Always use try-except** for migration operations
2. **Check return values** for success/failure
3. **Review logs** after migration
4. **Use dry-run mode** to preview changes
5. **Check reports** for detailed results

## Examples

### Simple Migration

```python
from jxls_migration_tool import JxlsMigrationTool

tool = JxlsMigrationTool(verbose=True)

# Migrate directory
success = tool.migrate_directory(
    input_dir="templates",
    output_dir="migrated",
    keep_extension=True
)

if success:
    tool.generate_report("migrated")
    print("Migration completed successfully")
```

### Custom Processing

```python
from jxls_migration_tool import JxlsMigrationTool, detect_excel_format

tool = JxlsMigrationTool(verbose=True)

# Process files with custom logic
files = ["file1.xls", "file2.xlsx", "file3.xls"]

for file in files:
    # Detect format first
    fmt = detect_excel_format(file)
    print(f"{file}: {fmt} format")

    # Migrate
    success = tool.migrate_file(file, keep_extension=True)

    if success:
        print(f"✅ {file} migrated")
    else:
        print(f"❌ {file} failed")
```

### Error Handling

```python
from jxls_migration_tool import JxlsMigrationTool

tool = JxlsMigrationTool(verbose=True)

try:
    success = tool.migrate_file("template.xls", "output.xls")

    if success:
        print("Migration successful")
    else:
        print("Migration failed - check migration report and logs")
        print("See: output/migration_report.md")

except FileNotFoundError:
    print("Error: Input file not found")
except PermissionError:
    print("Error: Permission denied - check output directory permissions")
except ValueError as e:
    print(f"Error: Invalid file format - {e}")
except Exception as e:
    print(f"Unexpected error: {e}")
    print("Check jxls_migration.log for details")
```

## Type Hints

The tool uses Python type hints for better IDE support and documentation.

### Import Types

```python
from typing import Optional, Dict, List
from pathlib import Path
```

### Common Type Aliases

```python
FilePath = str  # Path to file
DirPath = str   # Path to directory
Bool = bool     # Boolean
```

### Method Signatures with Types

```python
from typing import Optional, Dict, List, Union

class JxlsMigrationTool:
    def migrate_file(
        self,
        input_path: FilePath,
        output_path: Optional[FilePath] = None,
        keep_extension: Bool = True
    ) -> Bool:
        """Migrate a single file."""
        pass

    def migrate_directory(
        self,
        input_dir: DirPath,
        output_dir: DirPath,
        keep_extension: Bool = True,
        dry_run: Bool = False
    ) -> Bool:
        """Migrate directory."""
        pass

    def generate_markdown_report(
        self,
        output_dir: DirPath,
        stats: Dict[str, Union[int, str]]
    ) -> None:
        """Generate report."""
        pass
```

## Advanced Usage

### Custom Command Processor

```python
from jxls_migration_tool import JxlsCommand

class CustomCommand(JxlsCommand):
    """Custom JXLS instruction processor."""

    def parse(self) -> dict:
        # Parse custom instruction
        return {"type": "custom", "value": self.content}

    def convert(self) -> str:
        # Convert to JXLS 2.x format
        return f"custom:{self.content}"

# Register custom processor
tool = JxlsMigrationTool()
# ... (registration logic would go here)
```

---

For more examples, see the [examples directory](../examples/).

For questions or issues, please check the [troubleshooting section](USAGE.md#troubleshooting) or [open an issue](https://github.com/your-org/jxls-migration-tool/issues).
