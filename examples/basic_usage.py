#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Basic Usage Example for JXLS Migration Tool

This example demonstrates how to use the JXLS migration tool
programmatically in your Python code.

Version: 3.4 (Command Fix Update)
Changes:
  - Fixed jx:each comment generation issue
  - Fixed jx:area position error (now correctly added to A1)
  - Enhanced error handling for Excel format conversion
  - Robust migration with automatic fallback mechanism
  - Support for Excel files with incomplete format information
"""

import os
import sys
from pathlib import Path

# Add parent directory to path if running standalone
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from jxls_migration_tool import JxlsMigrationTool


def example_single_file_migration():
    """Example: Migrate a single file"""
    print("=" * 60)
    print("Example 1: Single File Migration")
    print("=" * 60)

    # Initialize the migration tool
    tool = JxlsMigrationTool(verbose=True)

    # Input and output file paths
    input_file = "examples/data/sample_template.xls"
    output_file = "examples/output/migrated_template.xls"

    # Create output directory if it doesn't exist
    os.makedirs(os.path.dirname(output_file), exist_ok=True)

    try:
        # Migrate the file
        result = tool.migrate_file(
            input_path=input_file,
            output_path=output_file,
            keep_extension=True
        )

        if result:
            print(f"✅ Successfully migrated: {input_file}")
            print(f"   Output: {output_file}")
        else:
            print(f"❌ Migration failed: {input_file}")

    except Exception as e:
        print(f"❌ Error: {e}")


def example_directory_migration():
    """Example: Migrate all files in a directory"""
    print("\n" + "=" * 60)
    print("Example 2: Directory Migration")
    print("=" * 60)

    # Initialize the migration tool
    tool = JxlsMigrationTool(verbose=True)

    # Input and output directories
    input_dir = "examples/data/templates"
    output_dir = "examples/output/templates"

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    try:
        # Migrate the directory
        result = tool.migrate_directory(
            input_dir=input_dir,
            output_dir=output_dir,
            keep_extension=True,
            dry_run=False
        )

        if result:
            print(f"✅ Successfully migrated directory: {input_dir}")
            print(f"   Output: {output_dir}")

            # Generate report
            tool.generate_report(output_dir)
            print(f"   Report: {output_dir}/migration_report.md")
        else:
            print(f"❌ Migration failed for directory: {input_dir}")

    except Exception as e:
        print(f"❌ Error: {e}")


def example_dry_run():
    """Example: Preview changes without modifying files"""
    print("\n" + "=" * 60)
    print("Example 3: Dry Run (Preview Changes)")
    print("=" * 60)

    # Initialize the migration tool
    tool = JxlsMigrationTool(verbose=True)

    # Input directory
    input_dir = "examples/data/templates"

    try:
        # Run in dry-run mode
        print(f"Scanning directory (dry-run): {input_dir}")
        print("-" * 60)

        result = tool.migrate_directory(
            input_dir=input_dir,
            output_dir="examples/output",
            keep_extension=True,
            dry_run=True
        )

        if result:
            print("\n✅ Dry run completed successfully")
            print("   (No files were modified)")
        else:
            print("\n❌ Dry run encountered issues")

    except Exception as e:
        print(f"❌ Error: {e}")


def example_custom_processing():
    """Example: Custom processing with file filtering"""
    print("\n" + "=" * 60)
    print("Example 4: Custom Processing")
    print("=" * 60)

    # Initialize the migration tool
    tool = JxlsMigrationTool(verbose=True)

    input_dir = "examples/data/templates"
    output_dir = "examples/output/filtered"

    # Create output directory
    os.makedirs(output_dir, exist_ok=True)

    # Get list of files
    files = list(Path(input_dir).glob("*.xls")) + list(Path(input_dir).glob("*.xlsx"))

    print(f"Found {len(files)} Excel files")
    print("-" * 60)

    success_count = 0
    fail_count = 0

    for file_path in files:
        try:
            # Custom logic: Only process files with "export" in the name
            if "export" in file_path.name.lower():
                output_path = Path(output_dir) / file_path.name

                print(f"Processing: {file_path.name}")

                result = tool.migrate_file(
                    input_path=str(file_path),
                    output_path=str(output_path),
                    keep_extension=True
                )

                if result:
                    success_count += 1
                    print(f"  ✅ Success")
                else:
                    fail_count += 1
                    print(f"  ❌ Failed")
            else:
                print(f"Skipping: {file_path.name} (no 'export' in name)")

        except Exception as e:
            fail_count += 1
            print(f"  ❌ Error: {e}")

    print("-" * 60)
    print(f"Results: {success_count} succeeded, {fail_count} failed")


def main():
    """Run all examples"""
    print("\n")
    print("╔" + "=" * 58 + "╗")
    print("║" + " " * 58 + "║")
    print("║" + "  JXLS Migration Tool - Python API Examples  ".center(58) + "║")
    print("║" + " " * 58 + "║")
    print("╚" + "=" * 58 + "╝")

    # Run examples
    example_single_file_migration()
    example_directory_migration()
    example_dry_run()
    example_custom_processing()

    print("\n" + "=" * 60)
    print("All examples completed!")
    print("=" * 60)


if __name__ == "__main__":
    main()
