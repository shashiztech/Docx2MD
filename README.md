# Docx2MDConverter VS Code Extension

## Features
- Convert DOCX files to Markdown with advanced formatting
- Preserves bold, italic, underline, lists, indentation
- Extracts and places images inline at correct positions
- Handles advanced tables and multi-column layouts
- Avoids unwanted markdown code blocks
- Customizable output directory per document

## Usage
1. Open the Command Palette (Ctrl+Shift+P)
2. Run `Docx2MD: Convert DOCX to Markdown`
3. Select a DOCX file and output directory
4. The extension will use the Python backend to convert and show results in the Output panel

## Requirements
- Python 3.x must be installed and available in PATH
- The Python script `docx_to_markdown_converter.py` must be present in the extension workspace or configured path
- Required Python packages: `python-docx`, `lxml`, etc. (see script requirements)

## Development
- Built with TypeScript and VS Code Extension API
- Backend conversion handled by Python script

## License
Copyright (c) 2025 Shashikanta Parida

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
