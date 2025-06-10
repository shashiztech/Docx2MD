# Docx2MD Converter

A powerful VS Code extension that converts DOCX and DOC files to Markdown with advanced formatting preservation, automatic image extraction, and intelligent table conversion.

## ✨ Features

- 🔄 **Universal Conversion**: Support for both DOCX and DOC files
- 🖼️ **Image Extraction**: Automatically extracts and organizes images
- 📝 **Formatting Preservation**: Maintains bold, italic, underline, and other formatting
- 📊 **Smart Table Conversion**: Converts tables with intelligent formatting detection
- 🔗 **Hyperlink Support**: Preserves hyperlinks with proper markdown syntax
- 📋 **Table of Contents**: Auto-generates TOC from document headings
- 📈 **Detailed Reports**: Comprehensive conversion reports with statistics
- 🎯 **Context Menu Integration**: Right-click conversion from file explorer
- ⚙️ **Configurable**: Customizable output directories and behavior
- 🐍 **Python Backend**: Uses robust Python script for reliable conversion

## 🔧 Requirements

### System Requirements
- **VS Code**: Version 1.74.0 or higher
- **Python**: Version 3.6 or higher (uses standard library only)
- **Operating System**: Windows, macOS, or Linux

### Python Compatibility Matrix
| Python Version | Status | Notes |
|---------------|---------|-------|
| 3.6 | ✅ Supported | Minimum version (f-strings, typing) |
| 3.7 | ✅ Supported | Full compatibility |
| 3.8 | ✅ Recommended | Better performance |
| 3.9+ | ✅ Supported | Latest features |

**No external Python packages required** - uses only standard library modules for maximum compatibility.

## 📦 Installation

### From VS Code Marketplace
1. Open VS Code
2. Go to Extensions (Ctrl+Shift+X)
3. Search for "Docx2MD Converter"
4. Click Install

### Manual Installation
1. Download the `.vsix` file from releases
2. Open VS Code
3. Run `Extensions: Install from VSIX...` command
4. Select the downloaded file

### Python Setup
Ensure Python 3.6+ is installed and accessible:
```bash
python --version    # Should show 3.6 or higher
# or
python3 --version   # Alternative command
```

## 🚀 Usage

### Method 1: Context Menu (Recommended)
1. Right-click any `.docx` or `.doc` file in VS Code Explorer
2. Select **"Convert DOCX/DOC to Markdown"**
3. Conversion starts automatically

### Method 2: Command Palette
1. Open Command Palette (`Ctrl+Shift+P` / `Cmd+Shift+P`)
2. Type **"Convert DOCX/DOC to Markdown"**
3. Select the command and choose your file

### Method 3: File Explorer
1. Select a `.docx` or `.doc` file in the explorer
2. Use the context menu or command palette

## 📁 Output Structure

The extension creates a well-organized output structure:

```
TargetMDDirectory/
└── document-name/
    ├── document-name.md          # Main markdown file
    ├── images/                   # Extracted images folder
    │   ├── image_001_chart.png
    │   ├── image_002_diagram.jpg
    │   └── ...
    └── conversion_report.md      # Detailed conversion report
```

## ⚙️ Configuration

Configure the extension through VS Code settings (`Ctrl+,`):

```json
{
    "docx2mdconverter.pythonPath": "python",
    "docx2mdconverter.outputDirectory": "TargetMDDirectory",
    "docx2mdconverter.showReport": true,
    "docx2mdconverter.autoOpenResult": true
}
```

### Available Settings

| Setting | Type | Default | Description |
|---------|------|---------|-------------|
| `pythonPath` | string | `"python"` | Path to Python executable |
| `outputDirectory` | string | `"TargetMDDirectory"` | Output folder name |
| `showReport` | boolean | `true` | Show conversion report |
| `autoOpenResult` | boolean | `true` | Auto-open converted file |

## 📊 Supported Features

| Feature | DOCX | DOC | Notes |
|---------|------|-----|-------|
| **Text Extraction** | ✅ Full | ✅ Full | Complete text content |
| **Formatting** | ✅ Full | ⚠️ Basic | Bold, italic, underline |
| **Images** | ✅ Full | ❌ Limited | Automatic extraction & organization |
| **Tables** | ✅ Smart | ⚠️ Basic | Intelligent formatting detection |
| **Hyperlinks** | ✅ Full | ❌ No | Complete URL preservation |
| **Headings** | ✅ Full | ⚠️ Heuristic | TOC generation |
| **Lists** | ✅ Full | ⚠️ Basic | Bullet and numbered lists |

## 🔍 Advanced Features

### Smart Table Detection
The converter intelligently detects and formats different table types:
- **Two-column layouts**: Key-value pairs, image-text combinations
- **Multi-category tables**: Complex data structures with categories
- **Standard tables**: Traditional markdown table format

### Image Management
- **Automatic extraction**: All embedded images extracted to `images/` folder
- **Intelligent naming**: Sequential numbering with descriptive names
- **Format preservation**: Maintains original image formats (PNG, JPG, etc.)
- **Markdown integration**: Proper markdown image syntax with relative paths

### Document Structure Analysis
- **Heading detection**: Multiple methods for identifying headings
- **TOC generation**: Automatic table of contents with anchor links
- **Style preservation**: Bold, italic, underline formatting maintained
- **List handling**: Proper markdown list formatting

## 🐛 Troubleshooting

### Common Issues

#### Python Not Found
```
❌ Python not found or not accessible
```
**Solutions:**
- Install Python 3.6+ from [python.org](https://python.org)
- Add Python to system PATH
- Configure `docx2mdconverter.pythonPath` setting

#### Script Not Found
```
❌ Python converter script not found
```
**Solutions:**
- Extension will offer to install the script automatically
- Manually place `docx_to_markdown_converter.py` in workspace
- Use "Locate Script" option from error dialog

#### Permission Denied
```
❌ Permission denied
```
**Solutions:**
- Check file permissions
- Ensure the document isn't open in another application
- Run VS Code as administrator (Windows)

#### Conversion Timeout
```
❌ Conversion timed out
```
**Solutions:**
- Try with smaller/simpler documents first
- Close other applications to free memory
- Check if document is corrupted

### Debugging Steps

1. **Check Output Panel**: View detailed logs in "Docx2MD Converter" output channel
2. **Verify Python**: Run `python --version` in terminal
3. **Test Script**: Manually run the Python script from command line
4. **Check File**: Ensure document isn't password-protected or corrupted

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guidelines](CONTRIBUTING.md) for details.

### Development Setup
1. Clone the repository
2. Run `npm install`
3. Open in VS Code
4. Press F5 to launch extension host

### Reporting Issues
Please include:
- VS Code version
- Python version
- Error messages from output panel
- Sample document (if possible)

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

## 🎯 Roadmap

- [ ] Batch conversion support
- [ ] Custom styling options
- [ ] Export to multiple formats
- [ ] Cloud storage integration
- [ ] Real-time preview
- [ ] Enhanced DOC support

## 📞 Support

- 🐛 **Issues**: [GitHub Issues](https://github.com/your-username/docx2md-converter/issues)
- 💬 **Discussions**: [GitHub Discussions](https://github.com/your-username/docx2md-converter/discussions)
- 📧 **Email**: your-email@example.com

## 🙏 Acknowledgments

- Python community for excellent standard library
- VS Code team for the extension API
- Contributors and users for feedback and improvements

---

## Following extension guidelines

Ensure that you've read through the extensions guidelines and follow the best practices for creating your extension.

* [Extension Guidelines](https://code.visualstudio.com/api/references/extension-guidelines)

## Working with Markdown

You can author your README using Visual Studio Code. Here are some useful editor keyboard shortcuts:

* Split the editor (`Cmd+\` on macOS or `Ctrl+\` on Windows and Linux).
* Toggle preview (`Shift+Cmd+V` on macOS or `Shift+Ctrl+V` on Windows and Linux).
* Press `Ctrl+Space` (Windows, Linux, macOS) to see a list of Markdown snippets.

## For more information

* [Visual Studio Code's Markdown Support](http://code.visualstudio.com/docs/languages/markdown)
* [Markdown Syntax Reference](https://help.github.com/articles/markdown-basics/)

**Enjoy!**
