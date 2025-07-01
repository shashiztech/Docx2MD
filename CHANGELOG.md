# Change Log

All notable changes to the "Docx2MD Converter" extension will be documented in this file.

## [0.1.4] - 2025-07-01

### 🛡️ Enhanced Error Handling & Sensitive Content Support
- **Advanced Error Recovery**: Comprehensive error handling system that gracefully handles document issues
- **Sensitive Content Detection**: Automatically detects and handles sensitivity labels, DRM, and protected content
- **Smart Issue Resolution**: Converts accessible content while safely handling restricted elements
- **Informative Messaging**: Replaced error messages with helpful handled issue notifications

### 🎯 User Experience Improvements
- **Handled vs Error Distinction**: Clear differentiation between handled issues and actual errors
- **Detailed Conversion Reports**: Enhanced reports with issue summaries and technical details
- **Success Despite Issues**: Shows success messages when conversion completes with handled issues
- **Progressive Error Handling**: Attempts multiple recovery strategies before failing

### 🔧 Technical Enhancements
- **ConversionWarning System**: New warning management system for tracking and reporting issues
- **Sensitivity Label Detection**: Checks for document protection indicators in DOCX files
- **Graceful Degradation**: Continues processing when encountering non-critical issues
- **Enhanced Logging**: Better logging and reporting of conversion processes and issues

### 🧹 Project Cleanup
- **Removed Test Files**: Cleaned up TP_KT.docx and test output directories
- **Production Ready**: Package now contains only production-ready files

## [0.1.3] - 2025-06-26

### ✨ New Features
- **Review & Feedback System**: Added intelligent review prompt that appears once per installation after 3 successful conversions
- **User Engagement**: Encourages marketplace ratings and feedback to help improve the extension
- **Smart Prompting**: Non-intrusive system that respects user choice and only asks once

### 🎯 User Experience
- **Friendly Messaging**: Emphasizes that this is a free extension and feedback helps development
- **Multiple Options**: Users can rate & review, provide feedback, or dismiss the prompt
- **Direct Links**: One-click access to marketplace review page and feedback channels

### 🔧 Technical Implementation
- **Usage Tracking**: Tracks conversion count using VS Code's global state
- **Privacy Friendly**: Only stores anonymous usage statistics locally
- **Error Handling**: Review system won't interfere with main conversion functionality

## [0.1.2] - 2025-06-10

### 🐛 Bug Fixes
- **Fixed Context Menu**: Corrected regex pattern for file extension detection to ensure "Convert DOCX/DOC to Markdown" appears in Explorer context menu
- **Enhanced Menu Accessibility**: Added context menu option to editor title bar for better accessibility
- **Improved File Detection**: Changed from regex pattern to explicit extension matching for more reliable detection

### 🔧 Technical Changes
- Updated `when` clause from regex `/\\.(docx|doc)$/i` to explicit `resourceExtname == '.docx' || resourceExtname == '.doc'`
- Added `editor/title/context` menu location for additional access points

## [0.1.1] - 2025-06-10

### 🎨 Visual Updates
- **Updated Extension Icon**: Changed to custom docx2markdown_skp icon for better visual identity
- **Icon Optimization**: Converted JPEG icon to PNG format and resized to standard 128x128 pixels

### 🔧 Maintenance
- Updated package version for new release

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