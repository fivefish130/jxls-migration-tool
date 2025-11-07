# Contributing to JXLS Migration Tool

Thank you for your interest in contributing to the JXLS Migration Tool! This document provides guidelines and information for contributors.

## Table of Contents

- [Code of Conduct](#code-of-conduct)
- [How to Contribute](#how-to-contribute)
- [Development Setup](#development-setup)
- [Coding Standards](#coding-standards)
- [Testing](#testing)
- [Documentation](#documentation)
- [Submitting Pull Requests](#submitting-pull-requests)
- [Issue Reporting](#issue-reporting)

## Code of Conduct

By participating in this project, you are expected to uphold our Code of Conduct:

- Be respectful and inclusive
- Focus on constructive feedback
- Be patient with new contributors
- Respect different viewpoints and experiences

## How to Contribute

We welcome various types of contributions:

1. 🐛 **Bug Reports** - Report bugs via GitHub Issues
2. ✨ **Feature Requests** - Suggest new features via GitHub Issues
3. 🔧 **Code Contributions** - Fix bugs or implement features
4. 📚 **Documentation** - Improve docs, examples, or this guide
5. 🧪 **Testing** - Add or improve test cases

## Development Setup

### Prerequisites

- Python 3.6 or higher
- Git
- Virtual environment tool (venv, conda, etc.)

### Setup

1. **Fork and clone the repository**
```bash
git clone https://github.com/your-username/jxls-migration-tool.git
cd jxls-migration-tool
```

2. **Create a virtual environment**
```bash
# Using venv
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Or using conda
conda create -n jxls-migration python=3.12
conda activate jxls-migration
```

3. **Install dependencies**
```bash
# Install production dependencies
pip install -r requirements.txt

# Install development dependencies
pip install -r requirements.txt
pip install pytest pytest-cov black flake8 mypy
```

4. **Verify installation**
```bash
python jxls_migration_tool.py --help
```

## Coding Standards

### Python Style Guide

We follow PEP 8 with these additions:

- **Line Length**: Maximum 88 characters (Black default)
- **Docstrings**: Use Google-style docstrings
- **Type Hints**: Use type hints for all public functions and methods
- **Imports**: Group imports (standard library, third-party, local)

### Code Formatting

Use **Black** for code formatting:

```bash
# Format code
black .

# Check formatting
black --check .
```

### Linting

Use **flake8** for linting:

```bash
# Run linter
flake8 .

# With plugins
flake8 jxls_migration_tool.py
```

### Type Checking

Use **mypy** for type checking:

```bash
# Check types
mypy jxls_migration_tool.py
```

## Testing

### Running Tests

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=jxls_migration_tool --cov-report=html

# Run specific test
pytest tests/test_migration.py::test_single_file
```

### Writing Tests

- Place tests in `tests/` directory
- Name test files `test_*.py`
- Name test functions `test_*`
- Use descriptive test names
- Include docstrings for test functions

Example test:

```python
def test_single_file_migration():
    """Test migration of a single file."""
    tool = JxlsMigrationTool(verbose=False)
    result = tool.migrate_file(
        "tests/data/sample.xls",
        "tests/output/migrated.xls"
    )
    assert result is True
```

### Test Coverage

Aim for >80% test coverage:

```bash
pytest --cov=jxls_migration_tool --cov-report=term-missing
```

## Documentation

### Code Documentation

- All public functions and classes must have docstrings
- Include type hints
- Add usage examples for complex functions

Example:

```python
def migrate_file(
    self,
    input_path: str,
    output_path: Optional[str] = None,
    keep_extension: bool = True
) -> bool:
    """
    Migrate a single Excel file.

    Args:
        input_path: Path to input Excel file.
        output_path: Path to output file. If None, uses input_path.
        keep_extension: Keep original file extension.

    Returns:
        True if migration successful, False otherwise.

    Example:
        >>> tool = JxlsMigrationTool()
        >>> tool.migrate_file("template.xls", "migrated.xls")
        True
    """
```

### User Documentation

- Update README.md for major changes
- Update USAGE.md for new features
- Update API.md for API changes
- Update CHANGELOG.md for all changes

## Submitting Pull Requests

### Before Submitting

1. **Run all tests**
```bash
pytest
```

2. **Check code formatting**
```bash
black --check .
flake8 .
mypy jxls_migration_tool.py
```

3. **Update documentation**
- Update CHANGELOG.md
- Update relevant docs
- Add/update examples

4. **Write clear commit messages**
```
feat: add support for jx:multiSheet instruction

- Implement MultiSheetCommand class
- Add multiSheet parsing and conversion
- Update documentation
- Add test cases
```

### PR Process

1. **Create a feature branch**
```bash
git checkout -b feature/your-feature-name
```

2. **Make your changes**
- Follow coding standards
- Add tests
- Update documentation

3. **Commit your changes**
```bash
git add .
git commit -m "feat: describe your change"
```

4. **Push to your fork**
```bash
git push origin feature/your-feature-name
```

5. **Create a Pull Request**
- Use the PR template
- Reference any related issues
- Add screenshots/examples if applicable

### PR Template

```markdown
## Description
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
- [ ] Tests pass locally
- [ ] Added/updated tests for changes
- [ ] Manual testing completed

## Checklist
- [ ] Code follows project style guidelines
- [ ] Self-review completed
- [ ] Documentation updated
- [ ] No new warnings introduced
```

## Issue Reporting

### Bug Reports

Include the following information:

- **Description**: Clear description of the bug
- **Steps to Reproduce**: Detailed steps
- **Expected Behavior**: What should happen
- **Actual Behavior**: What actually happens
- **Environment**: OS, Python version, tool version
- **Error Message**: Full error message and stack trace
- **Sample File**: Minimal example file (if possible)

### Feature Requests

Include the following information:

- **Problem Statement**: What problem does this solve?
- **Proposed Solution**: How should this work?
- **Alternatives**: Any alternative solutions considered?
- **Additional Context**: Any other relevant information

## Release Process

1. Update version in `setup.py`
2. Update CHANGELOG.md
3. Create release PR
4. Tag release after merge
5. Build and publish to PyPI (maintainers only)

## Questions?

Feel free to:

- Open an issue for questions
- Start a discussion for general topics
- Review existing issues and PRs

## Recognition

Contributors will be recognized in:
- CHANGELOG.md
- README.md (for significant contributions)
- GitHub contributors list

Thank you for contributing! 🎉
