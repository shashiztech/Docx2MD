# Extension Testing Guide

## Testing the Docx2MD Converter Extension

### Prerequisites
1. Ensure Python 3.6+ is installed and accessible via `python` command
2. Have a sample DOCX or DOC file ready for testing
3. VS Code with the extension installed

### Test Scenarios

#### 1. Basic Conversion Test
- Right-click on a DOCX file in VS Code Explorer
- Select "Convert DOCX/DOC to Markdown"
- Verify conversion completes successfully
- Check output files are created in TargetMDDirectory

#### 2. Command Palette Test
- Open Command Palette (Ctrl+Shift+P)
- Type "Convert DOCX/DOC to Markdown"
- Select a test file
- Verify conversion process

#### 3. Error Handling Test
- Test with non-existent file
- Test with corrupted document
- Test with wrong file extension
- Verify appropriate error messages

#### 4. Configuration Test
- Change settings in VS Code preferences
- Test different Python paths
- Test different output directories
- Verify settings are respected

### Expected Outputs

#### File Structure
```
TargetMDDirectory/
└── document-name/
    ├── document-name.md
    ├── images/
    │   └── (extracted images)
    └── conversion_report.md
```

#### Features to Verify
- [x] Text conversion accuracy
- [x] Image extraction and placement
- [x] Table formatting
- [x] Heading detection
- [x] Hyperlink preservation
- [x] List formatting
- [x] Error handling
- [x] Progress indication
- [x] Output organization

### Manual Testing Checklist

- [ ] Extension activates without errors
- [ ] Python version detection works
- [ ] Script discovery functions correctly
- [ ] File selection dialog appears
- [ ] Progress notification shows
- [ ] Conversion completes successfully
- [ ] Output files are created
- [ ] Markdown content is accurate
- [ ] Images are extracted correctly
- [ ] Report generation works
- [ ] Error messages are helpful
- [ ] Settings configuration works

### Common Issues and Solutions

#### Python Not Found
- Ensure Python is in PATH
- Configure custom Python path in settings
- Try using `python3` command

#### Script Not Found
- Copy script to workspace manually
- Use "Locate Script" option
- Check extension installation

#### Permission Errors
- Close document in other applications
- Run VS Code as administrator
- Check file permissions

### Performance Testing

#### Small Documents (< 1MB)
- Should complete in under 10 seconds
- Memory usage should be minimal

#### Medium Documents (1-10MB)
- Should complete in under 30 seconds
- Progress indication should be accurate

#### Large Documents (> 10MB)
- May take 1-2 minutes
- Should not timeout
- Progress should update regularly

### Reporting Issues

When reporting issues, include:
- VS Code version
- Python version
- Operating system
- Error messages from output panel
- Sample document (if possible)
- Steps to reproduce

### Success Criteria

The extension is ready for publication when:
- All test scenarios pass
- Error handling is comprehensive
- Performance is acceptable
- Documentation is complete
- Code quality is high
- No critical bugs remain
