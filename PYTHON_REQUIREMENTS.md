# Python Requirements for Docx2MD Converter

## Minimum Python Version
- **Required**: Python 3.6 or higher
- **Recommended**: Python 3.8 or higher

## Backward Compatibility

The converter is designed for maximum backward compatibility and uses only Python standard library modules:

### Python 3.6+ Features Used:
- f-string formatting (Python 3.6+)
- Type hints with `typing` module
- `pathlib.Path` for cross-platform file operations
- Enhanced `zipfile` and `xml.etree.ElementTree` modules

### Standard Library Dependencies:
- `os`, `sys` - System operations
- `zipfile` - DOCX file handling (ZIP format)
- `re` - Regular expressions for text processing
- `pathlib` - Modern path handling
- `typing` - Type hints for better code quality
- `xml.etree.ElementTree` - XML parsing for DOCX content
- `logging` - Structured logging
- `struct` - Binary data handling for DOC files

## No External Dependencies Required

The converter intentionally uses only standard library modules to ensure:
- ✅ Easy installation and deployment
- ✅ No dependency management issues
- ✅ Maximum compatibility across Python versions
- ✅ Reduced security concerns
- ✅ Faster startup times

## Optional Enhancements

For enhanced .doc file support, users can optionally install:
```bash
pip install python-docx>=0.8.10  # Enhanced DOCX processing
pip install olefile>=0.46        # Better DOC file support
```

However, these are **not required** - the converter will work with just the standard library.

## Installation Verification

To verify your Python installation is compatible:

```bash
python --version          # Should show 3.6 or higher
python3 --version         # Alternative command
```

## VS Code Extension Integration

The VS Code extension automatically:
- Detects Python installation
- Verifies version compatibility (3.6+)
- Falls back to `python3` command if `python` fails
- Provides clear error messages for version issues
- Supports custom Python path configuration
