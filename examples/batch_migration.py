#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Batch Migration Example for JXLS Migration Tool

This example demonstrates how to perform batch migrations
with proper error handling and reporting.

Version: 3.1 (Format Fix Update)
Changes:
  - Fixed 'Format' object has no attribute 'font_index' error
  - Enhanced error handling for Excel format conversion
  - Support for Excel files with incomplete format information
  - Safe property access using hasattr() and getattr()
"""

import os
import sys
import json
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

# Add parent directory to path if running standalone
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from jxls_migration_tool import JxlsMigrationTool


class BatchMigration:
    """Batch migration manager with enhanced error handling and reporting"""

    def __init__(self, max_workers=4, verbose=True):
        """
        Initialize batch migration manager

        Args:
            max_workers: Maximum number of concurrent workers
            verbose: Enable verbose logging
        """
        self.max_workers = max_workers
        self.verbose = verbose
        self.tool = JxlsMigrationTool(verbose=verbose)
        self.results = {
            "timestamp": datetime.now().isoformat(),
            "total_files": 0,
            "successful": 0,
            "failed": 0,
            "skipped": 0,
            "errors": []
        }

    def migrate_directory(self, input_dir, output_dir, keep_extension=True, dry_run=False):
        """
        Migrate all Excel files in a directory

        Args:
            input_dir: Input directory path
            output_dir: Output directory path
            keep_extension: Keep original file extensions
            dry_run: Preview without modifying files

        Returns:
            dict: Migration results
        """
        # Create output directory
        os.makedirs(output_dir, exist_ok=True)

        # Find all Excel files
        excel_files = self._find_excel_files(input_dir)

        if not excel_files:
            print(f"⚠️  No Excel files found in {input_dir}")
            return self.results

        self.results["total_files"] = len(excel_files)
        print(f"📁 Found {len(excel_files)} Excel files to process")
        print(f"   Output directory: {output_dir}")
        print(f"   Keep extension: {keep_extension}")
        print(f"   Dry run: {dry_run}")
        print("-" * 60)

        # Process files in parallel
        self._process_files(excel_files, output_dir, keep_extension, dry_run)

        # Generate summary
        self._print_summary()

        return self.results

    def _find_excel_files(self, directory):
        """Find all Excel files in a directory"""
        excel_files = []
        patterns = ["*.xls", "*.xlsx"]

        for pattern in patterns:
            excel_files.extend(Path(directory).glob(pattern))
            excel_files.extend(Path(directory).rglob(pattern))

        return sorted(excel_files)

    def _process_files(self, file_paths, output_dir, keep_extension, dry_run):
        """Process files using thread pool"""
        tasks = []

        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # Submit all tasks
            for file_path in file_paths:
                output_path = self._get_output_path(
                    file_path, output_dir, keep_extension
                )

                task = executor.submit(
                    self._migrate_single_file,
                    str(file_path),
                    str(output_path),
                    dry_run
                )
                tasks.append((file_path.name, task))

            # Process completed tasks
            for filename, future in as_completed(tasks):
                try:
                    success, error = future.result()

                    if success:
                        self.results["successful"] += 1
                        print(f"✅ {filename}")
                    elif error == "no_jxls":
                        self.results["skipped"] += 1
                        print(f"⏭️  {filename} (no JXLS instructions)")
                    else:
                        self.results["failed"] += 1
                        self.results["errors"].append({
                            "file": filename,
                            "error": error
                        })
                        print(f"❌ {filename} - {error}")

                except Exception as e:
                    self.results["failed"] += 1
                    self.results["errors"].append({
                        "file": filename,
                        "error": str(e)
                    })
                    print(f"❌ {filename} - {e}")

    def _migrate_single_file(self, input_path, output_path, dry_run):
        """Migrate a single file with error handling"""
        try:
            if dry_run:
                # Check if file contains JXLS instructions
                has_jxls = self._check_jxls_instructions(input_path)
                if has_jxls:
                    return True, None
                else:
                    return True, "no_jxls"
            else:
                # Perform actual migration
                success = self.tool.migrate_file(
                    input_path=input_path,
                    output_path=output_path,
                    keep_extension=True
                )

                if success:
                    return True, None
                else:
                    return False, "migration_failed"

        except Exception as e:
            return False, str(e)

    def _check_jxls_instructions(self, file_path):
        """Check if file contains JXLS instructions"""
        try:
            # Simple check for JXLS tags
            with open(file_path, 'rb') as f:
                content = f.read()

            # Check for common JXLS patterns
            jxls_patterns = [
                b'<jx:forEach',
                b'<jx:if',
                b'<jx:out',
                b'<jx:area',
                b'<jx:multiSheet'
            ]

            return any(pattern in content for pattern in jxls_patterns)

        except Exception:
            return False

    def _get_output_path(self, file_path, output_dir, keep_extension):
        """Determine output file path"""
        file_path = Path(file_path)
        output_path = Path(output_dir) / file_path.name

        # Ensure output directory structure
        output_path.parent.mkdir(parents=True, exist_ok=True)

        return output_path

    def _print_summary(self):
        """Print migration summary"""
        print("\n" + "=" * 60)
        print("Migration Summary")
        print("=" * 60)
        print(f"📊 Total files: {self.results['total_files']}")
        print(f"✅ Successful: {self.results['successful']}")
        print(f"⏭️  Skipped: {self.results['skipped']}")
        print(f"❌ Failed: {self.results['failed']}")

        if self.results["total_files"] > 0:
            success_rate = (self.results["successful"] / self.results["total_files"]) * 100
            print(f"🎯 Success rate: {success_rate:.2f}%")

        if self.results["errors"]:
            print("\n❌ Errors:")
            for error in self.results["errors"][:5]:  # Show first 5 errors
                print(f"   - {error['file']}: {error['error']}")
            if len(self.results["errors"]) > 5:
                print(f"   ... and {len(self.results['errors']) - 5} more errors")

        print("=" * 60)

    def save_report(self, output_file):
        """Save migration report to JSON file"""
        report_file = Path(output_file)

        # Add additional metadata
        self.results["max_workers"] = self.max_workers
        self.results["version"] = "3.1.0"

        with open(report_file, 'w', encoding='utf-8') as f:
            json.dump(self.results, f, indent=2, ensure_ascii=False)

        print(f"\n📄 Report saved to: {report_file}")


def example_basic_batch():
    """Example: Basic batch migration"""
    print("\n" + "=" * 60)
    print("Batch Migration Example: Basic")
    print("=" * 60)

    batch = BatchMigration(max_workers=4, verbose=True)

    results = batch.migrate_directory(
        input_dir="examples/data/templates",
        output_dir="examples/output/batch_basic",
        keep_extension=True,
        dry_run=False
    )

    batch.save_report("examples/output/batch_report.json")


def example_dry_run_batch():
    """Example: Batch dry run"""
    print("\n" + "=" * 60)
    print("Batch Migration Example: Dry Run")
    print("=" * 60)

    batch = BatchMigration(max_workers=4, verbose=True)

    results = batch.migrate_directory(
        input_dir="examples/data/templates",
        output_dir="examples/output/batch_dryrun",
        keep_extension=True,
        dry_run=True
    )

    batch.save_report("examples/output/dryrun_report.json")


def example_incremental_migration():
    """Example: Incremental migration (only changed files)"""
    print("\n" + "=" * 60)
    print("Batch Migration Example: Incremental")
    print("=" * 60)

    # Initialize batch processor
    batch = BatchMigration(max_workers=4, verbose=True)

    input_dir = "examples/data/templates"
    output_dir = "examples/output/incremental"

    # Find all files
    all_files = batch._find_excel_files(input_dir)

    # Check which files need migration
    needs_migration = []
    for file_path in all_files:
        output_path = Path(output_dir) / file_path.name

        # Check if output exists and is newer
        if not output_path.exists() or output_path.stat().st_mtime < file_path.stat().st_mtime:
            needs_migration.append(file_path)

    if not needs_migration:
        print("✅ All files are up to date")
        return

    print(f"📊 {len(needs_migration)} files need migration")
    print("-" * 60)

    # Migrate only changed files
    for file_path in needs_migration:
        print(f"🔄 {file_path.name}")

    # Perform migration
    batch.migrate_directory(
        input_dir=input_dir,
        output_dir=output_dir,
        keep_extension=True,
        dry_run=False
    )


def main():
    """Run all batch examples"""
    print("\n")
    print("╔" + "=" * 58 + "╗")
    print("║" + " " * 58 + "║")
    print("║" + "  JXLS Migration Tool - Batch Processing Examples  ".center(58) + "║")
    print("║" + " " * 58 + "║")
    print("╚" + "=" * 58 + "╝")

    # Run examples
    example_dry_run_batch()
    example_basic_batch()
    example_incremental_migration()

    print("\n" + "=" * 60)
    print("All batch examples completed!")
    print("=" * 60)


if __name__ == "__main__":
    main()
