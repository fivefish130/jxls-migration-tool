#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""Setup script for JXLS Migration Tool."""

from setuptools import setup, find_packages
import os

# Read README for long description
def read_readme():
    """Read README file for long description."""
    readme_path = os.path.join(os.path.dirname(__file__), 'README.md')
    if os.path.exists(readme_path):
        with open(readme_path, 'r', encoding='utf-8') as f:
            return f.read()
    return ""

# Read version from tool
def get_version():
    """Get version from the tool file."""
    tool_path = os.path.join(os.path.dirname(__file__), 'jxls_migration_tool.py')
    with open(tool_path, 'r', encoding='utf-8') as f:
        for line in f:
            if line.strip().startswith('版本:'):
                # Extract version from line like "版本: 3.0"
                return line.split(':')[1].strip().split()[0]
    return "3.1.0"

setup(
    name="jxls-migration-tool",
    version=get_version(),
    author="Claude Code",
    author_email="claude@anthropic.com",
    description="Production-ready tool for automated migration from JXLS 1.x to JXLS 2.14.0",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    url="https://github.com/your-org/jxls-migration-tool",
    project_urls={
        "Bug Tracker": "https://github.com/your-org/jxls-migration-tool/issues",
        "Documentation": "https://github.com/your-org/jxls-migration-tool/blob/main/README.md",
        "Source Code": "https://github.com/your-org/jxls-migration-tool",
    },
    packages=find_packages(),
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Intended Audience :: System Administrators",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Operating System :: Microsoft :: Windows :: Windows 10",
        "Operating System :: Microsoft :: Windows :: Windows 11",
        "Operating System :: POSIX :: Linux",
        "Operating System :: MacOS",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "Topic :: Text Processing :: General",
        "Topic :: Utilities",
    ],
    keywords=[
        "jxls",
        "excel",
        "migration",
        "templating",
        "spreadsheet",
        "xls",
        "xlsx",
        "automation",
    ],
    python_requires=">=3.6",
    install_requires=[
        "xlrd==2.0.1",
        "openpyxl>=3.0.0",
    ],
    extras_require={
        "dev": [
            "pytest>=6.0",
            "pytest-cov>=2.0",
            "black>=21.0",
            "flake8>=3.8",
            "mypy>=0.800",
        ],
        "test": [
            "pytest>=6.0",
            "pytest-cov>=2.0",
        ],
    },
    entry_points={
        "console_scripts": [
            "jxls-migrate=jxls_migration_tool:main",
        ],
    },
    include_package_data=True,
    package_data={
        "": [
            "docs/*.md",
            "examples/*.py",
            "LICENSE",
        ],
    },
    zip_safe=False,
    platforms="any",
)
