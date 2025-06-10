# Change Log

All notable changes to the "Docx2MD Converter" extension will be documented in this file.

## [0.1.0] - 2025-06-10

### 🚀 Initial Release

#### ✨ Features
- **Universal Document Conversion**: Support for both DOCX and DOC files
- **Intelligent Image Extraction**: Automatically extracts and organizes embedded images
- **Advanced Formatting Preservation**: Maintains bold, italic, underline, and complex formatting
- **Smart Table Conversion**: Intelligent detection and formatting of different table types
  - Two-column layouts (key-value pairs, image-text combinations)
  - Multi-category tables with proper categorization
  - Standard markdown table format
- **Hyperlink Support**: Preserves hyperlinks with proper markdown syntax
- **Automatic TOC Generation**: Creates table of contents from document headings
- **Context Menu Integration**: Right-click conversion from VS Code explorer
- **Progress Indication**: Real-time conversion progress with cancellation support
- **Comprehensive Error Handling**: Detailed error messages and troubleshooting guidance

#### 🔧 Technical Features
- **Python Version Compatibility**: Supports Python 3.6+ with automatic detection
- **Standard Library Only**: No external dependencies required for maximum compatibility
- **Cross-platform Support**: Works on Windows, macOS, and Linux
- **Configurable Settings**: Customizable Python path, output directory, and behavior
- **Auto-script Discovery**: Automatically finds Python converter script in multiple locations
- **Detailed Logging**: Comprehensive output logging for debugging and monitoring

#### 📁 Output Structure
- Organized output with separate directories per document
- Images extracted to dedicated `images/` folder
- Conversion reports with detailed statistics
- Automatic file opening and folder revelation

#### ⚙️ Configuration Options
- `docx2mdconverter.pythonPath`: Custom Python executable path
- `docx2mdconverter.outputDirectory`: Default output directory name
- `docx2mdconverter.showReport`: Toggle conversion report display
- `docx2mdconverter.autoOpenResult`: Auto-open converted files

#### 🐍 Python Backend
- Uses robust Python script with advanced document parsing
- Handles complex DOCX structure with XML namespace support
- Basic DOC file support with heuristic text extraction
- Intelligent heading detection and document structure analysis
- Advanced table processing with multiple formatting strategies

### 📊 Supported Features Matrix

| Feature | DOCX | DOC | Notes |
|---------|------|-----|-------|
| Text Extraction | ✅ Full | ✅ Full | Complete content preservation |
| Formatting | ✅ Full | ⚠️ Basic | Bold, italic, underline support |
| Images | ✅ Full | ❌ Limited | Automatic extraction & organization |
| Tables | ✅ Smart | ⚠️ Basic | Intelligent formatting detection |
| Hyperlinks | ✅ Full | ❌ No | Complete URL preservation |
| Headings | ✅ Full | ⚠️ Heuristic | TOC generation support |
| Lists | ✅ Full | ⚠️ Basic | Bullet and numbered lists |

### 🎯 Usage Methods
1. **Context Menu**: Right-click on DOCX/DOC files
2. **Command Palette**: Use "Convert DOCX/DOC to Markdown" command
3. **File Selection**: Automatic file picker when no file selected

### 🔍 Advanced Capabilities
- **Smart Image Handling**: Sequential numbering with descriptive names
- **Document Structure Analysis**: Multiple heading detection methods
- **List Processing**: Proper markdown list formatting
- **Error Recovery**: Graceful handling of corrupted or complex documents
- **Performance Optimization**: Efficient processing with progress feedback

### Known Limitations
- DOC file support is basic (heuristic text extraction)
- Complex DOC tables may not convert perfectly
- Password-protected documents are not supported
- Very large documents may require increased timeout settings

### Requirements
- VS Code 1.74.0 or higher
- Python 3.6 or higher (standard library only)
- Windows, macOS, or Linux operating system

---

## Future Roadmap

### Planned Features
- [ ] Batch conversion support for multiple files
- [ ] Custom styling and formatting options
- [ ] Export to additional formats (HTML, PDF)
- [ ] Cloud storage integration (OneDrive, Google Drive)
- [ ] Real-time preview functionality
- [ ] Enhanced DOC file support with better parsing
- [ ] Custom template support
- [ ] Markdown linting and validation
- [ ] Integration with documentation tools

### Performance Improvements
- [ ] Streaming conversion for large files
- [ ] Parallel processing for batch operations
- [ ] Memory optimization for large documents
- [ ] Caching for repeated conversions

---

## Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for details on:
- Setting up the development environment
- Code style and conventions
- Testing procedures
- Submitting pull requests

## Support

- 🐛 **Report Issues**: [GitHub Issues](https://github.com/your-username/docx2md-converter/issues)
- 💬 **Discussions**: [GitHub Discussions](https://github.com/your-username/docx2md-converter/discussions)
- 📧 **Contact**: your-email@example.com

---

**Thank you for using Docx2MD Converter!**