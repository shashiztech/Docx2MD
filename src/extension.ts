import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { execFile, spawn } from 'child_process';

let outputChannel: vscode.OutputChannel;

export function activate(context: vscode.ExtensionContext) {
    console.log('Docx2MD Converter extension is now active!');
    
    // Create output channel
    outputChannel = vscode.window.createOutputChannel('Docx2MD Converter');
    
    // Register the conversion command
    let convertDisposable = vscode.commands.registerCommand('docx2mdconverter.convertDocxToMarkdown', async (uri?: vscode.Uri) => {
        try {
            // Get the file path
            let filePath: string;
            
            if (uri) {
                filePath = uri.fsPath;
            } else {
                // Show file picker if no file selected
                const fileUri = await vscode.window.showOpenDialog({
                    canSelectFiles: true,
                    canSelectFolders: false,
                    canSelectMany: false,
                    filters: {
                        'Word Documents': ['docx', 'doc']
                    },
                    title: 'Select DOCX or DOC file to convert',
                    openLabel: 'Convert to Markdown'
                });
                
                if (!fileUri || fileUri.length === 0) {
                    return;
                }
                filePath = fileUri[0].fsPath;
            }
            
            // Validate file exists and has correct extension
            if (!fs.existsSync(filePath)) {
                vscode.window.showErrorMessage('Selected file does not exist.');
                return;
            }
            
            const ext = path.extname(filePath).toLowerCase();
            if (ext !== '.docx' && ext !== '.doc') {
                vscode.window.showErrorMessage('Please select a .docx or .doc file.');
                return;
            }
            
            // Get configuration
            const config = vscode.workspace.getConfiguration('docx2mdconverter');
            const pythonPath = config.get<string>('pythonPath', 'python');
            const outputDir = config.get<string>('outputDirectory', 'TargetMDDirectory');
            const showReport = config.get<boolean>('showReport', true);
            const autoOpenResult = config.get<boolean>('autoOpenResult', true);
            
            // Find Python script
            const scriptPath = await findPythonScript(context.extensionPath);
            if (!scriptPath) {
                const installScript = 'Install Python Script';
                const locateScript = 'Locate Script';
                
                const choice = await vscode.window.showErrorMessage(
                    'Python converter script (docx_to_markdown_converter.py) not found. Please ensure it\'s available in your workspace or extension directory.',
                    installScript,
                    locateScript
                );
                
                if (choice === installScript) {
                    await copyScriptToWorkspace(context.extensionPath);
                } else if (choice === locateScript) {
                    const scriptUri = await vscode.window.showOpenDialog({
                        canSelectFiles: true,
                        canSelectFolders: false,
                        canSelectMany: false,
                        filters: {
                            'Python Scripts': ['py']
                        },
                        title: 'Locate docx_to_markdown_converter.py'
                    });
                    
                    if (scriptUri && scriptUri.length > 0) {
                        // Copy to workspace
                        const workspaceFolder = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
                        if (workspaceFolder) {
                            const targetPath = path.join(workspaceFolder, 'docx_to_markdown_converter.py');
                            fs.copyFileSync(scriptUri[0].fsPath, targetPath);
                            vscode.window.showInformationMessage('Python script copied to workspace.');
                        }
                    }
                }
                return;
            }
            
            // Check Python version compatibility
            const pythonInfo = await checkPythonCompatibility(pythonPath);
            if (!pythonInfo.isCompatible) {
                const message = pythonInfo.version 
                    ? `Python ${pythonInfo.version} found, but Python 3.6+ is required.`
                    : 'Python not found or not accessible. Please install Python 3.6+ and ensure it\'s in your PATH.';
                
                vscode.window.showErrorMessage(message + ' Configure the Python path in settings if needed.');
                return;
            }
            
            outputChannel.show(true);
            outputChannel.appendLine(`🚀 Starting conversion of: ${path.basename(filePath)}`);
            outputChannel.appendLine(`🐍 Python version: ${pythonInfo.version}`);
            outputChannel.appendLine(`📄 Using script: ${scriptPath}`);
            outputChannel.appendLine(`📁 File type: ${ext.substring(1).toUpperCase()}`);
            
            // Determine output path
            const workspaceFolder = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath || path.dirname(filePath);
            const outputPath = path.join(workspaceFolder, outputDir);
            
            // Show progress
            await vscode.window.withProgress({
                location: vscode.ProgressLocation.Notification,
                title: `Converting ${path.basename(filePath)} to Markdown...`,
                cancellable: true
            }, async (progress, token) => {
                return new Promise<void>((resolve, reject) => {
                    // Execute Python script with arguments
                    const args = [scriptPath, filePath, '--output', outputPath];
                    
                    outputChannel.appendLine(`💻 Executing: ${pythonPath} ${args.join(' ')}`);
                    
                    const pythonProcess = execFile(pythonPath, args, {
                        cwd: workspaceFolder,
                        timeout: 120000, // 2 minute timeout
                        maxBuffer: 1024 * 1024 * 10 // 10MB buffer
                    }, (error, stdout, stderr) => {
                        if (token.isCancellationRequested) {
                            outputChannel.appendLine('❌ Conversion cancelled by user');
                            reject(new Error('Cancelled'));
                            return;
                        }
                        
                        if (error) {
                            outputChannel.appendLine(`❌ Error: ${error.message}`);
                            
                            let errorMessage = 'Conversion failed. Check the output panel for details.';
                            
                            if (error.code === 'ENOENT') {
                                errorMessage = 'Python executable not found. Please check your Python installation and PATH configuration.';
                            } else if (error.code === 'ETIMEDOUT') {
                                errorMessage = 'Conversion timed out. The file might be too large or complex.';
                            } else if (error.message.includes('Permission denied')) {
                                errorMessage = 'Permission denied. Please check file permissions.';
                            }
                            
                            vscode.window.showErrorMessage(errorMessage);
                            reject(error);
                            return;
                        }
                        
                        if (stderr) {
                            outputChannel.appendLine(`⚠️ stderr: ${stderr}`);
                        }
                        
                        if (stdout) {
                            outputChannel.appendLine(stdout);
                        }
                        
                        outputChannel.appendLine('✅ Conversion completed successfully!');
                        
                        // Show success message with options
                        handleConversionSuccess(filePath, outputPath, showReport, autoOpenResult);
                        resolve();
                    });
                    
                    // Handle cancellation
                    token.onCancellationRequested(() => {
                        pythonProcess.kill();
                    });
                    
                    // Handle process output in real-time
                    pythonProcess.stdout?.on('data', (data) => {
                        const output = data.toString();
                        outputChannel.append(output);
                        
                        // Update progress based on output
                        if (output.includes('Extracting images')) {
                            progress.report({ message: 'Extracting images...' });
                        } else if (output.includes('Converting text')) {
                            progress.report({ message: 'Converting text...' });
                        } else if (output.includes('Processing tables')) {
                            progress.report({ message: 'Processing tables...' });
                        }
                    });
                    
                    pythonProcess.stderr?.on('data', (data) => {
                        outputChannel.append(`stderr: ${data.toString()}`);
                    });
                });
            });
            
        } catch (error) {
            outputChannel.appendLine(`💥 Unexpected error: ${error}`);
            vscode.window.showErrorMessage(`Conversion failed: ${error}`);
        }
    });
    
    context.subscriptions.push(convertDisposable);
    context.subscriptions.push(outputChannel);
}

async function findPythonScript(extensionPath: string): Promise<string | null> {
    // Priority order for finding the script
    const searchPaths = [
        // 1. Extension directory (bundled with extension)
        path.join(extensionPath, 'docx_to_markdown_converter.py'),
        
        // 2. Workspace folders
        ...(vscode.workspace.workspaceFolders?.map(folder => 
            path.join(folder.uri.fsPath, 'docx_to_markdown_converter.py')
        ) || []),
        
        // 3. Current file directory (if a file is open)
        ...(vscode.window.activeTextEditor ? [
            path.join(path.dirname(vscode.window.activeTextEditor.document.fileName), 'docx_to_markdown_converter.py')
        ] : [])
    ];
    
    for (const scriptPath of searchPaths) {
        if (fs.existsSync(scriptPath)) {
            return scriptPath;
        }
    }
    
    return null;
}

interface PythonInfo {
    isCompatible: boolean;
    version?: string;
    executable?: string;
}

async function checkPythonCompatibility(pythonPath: string): Promise<PythonInfo> {
    // Try different Python executables
    const pythonCommands = [pythonPath, 'python3', 'python'];
    
    for (const cmd of pythonCommands) {
        try {
            const result = await new Promise<PythonInfo>((resolve) => {
                execFile(cmd, ['--version'], { timeout: 5000 }, (error, stdout, stderr) => {
                    if (error) {
                        resolve({ isCompatible: false });
                        return;
                    }
                    
                    const versionOutput = (stdout || stderr).trim();
                    const versionMatch = versionOutput.match(/Python (\d+)\.(\d+)\.(\d+)/);
                    
                    if (versionMatch) {
                        const major = parseInt(versionMatch[1]);
                        const minor = parseInt(versionMatch[2]);
                        const patch = parseInt(versionMatch[3]);
                        
                        // Check for Python 3.6+ (minimum for f-strings and typing)
                        const isCompatible = major === 3 && minor >= 6;
                        
                        resolve({
                            isCompatible,
                            version: `${major}.${minor}.${patch}`,
                            executable: cmd
                        });
                    } else {
                        resolve({ isCompatible: false });
                    }
                });
            });
            
            if (result.isCompatible) {
                return result;
            }
        } catch (error) {
            // Continue to next command
        }
    }
    
    return { isCompatible: false };
}

async function copyScriptToWorkspace(extensionPath: string): Promise<void> {
    const workspaceFolder = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
    if (!workspaceFolder) {
        vscode.window.showErrorMessage('No workspace folder open. Please open a folder first.');
        return;
    }
    
    const sourcePath = path.join(extensionPath, 'docx_to_markdown_converter.py');
    const targetPath = path.join(workspaceFolder, 'docx_to_markdown_converter.py');
    
    if (fs.existsSync(sourcePath)) {
        fs.copyFileSync(sourcePath, targetPath);
        vscode.window.showInformationMessage('Python script installed to workspace.');
    } else {
        vscode.window.showErrorMessage('Python script not found in extension. Please download it manually.');
    }
}

async function handleConversionSuccess(
    filePath: string,
    outputPath: string,
    showReport: boolean,
    autoOpenResult: boolean
): Promise<void> {
    const docName = path.basename(filePath, path.extname(filePath));
    const docOutputPath = path.join(outputPath, docName);
    const markdownFile = path.join(docOutputPath, `${docName}.md`);
    const reportFile = path.join(docOutputPath, 'conversion_report.md');
    
    // Show success notification with actions
    const actions = ['Open Output Folder', 'Open Markdown File'];
    if (fs.existsSync(reportFile)) {
        actions.push('View Report');
    }
    
    const message = `✅ Conversion completed! Document converted to Markdown successfully.`;
    const selection = await vscode.window.showInformationMessage(message, ...actions);
    
    switch (selection) {
        case 'Open Output Folder':
            vscode.commands.executeCommand('revealFileInOS', vscode.Uri.file(docOutputPath));
            break;
            
        case 'Open Markdown File':
            if (fs.existsSync(markdownFile)) {
                vscode.commands.executeCommand('vscode.open', vscode.Uri.file(markdownFile));
            } else {
                vscode.window.showErrorMessage('Markdown file not found.');
            }
            break;
            
        case 'View Report':
            vscode.commands.executeCommand('vscode.open', vscode.Uri.file(reportFile));
            break;
    }
    
    // Auto-open result if enabled
    if (autoOpenResult && fs.existsSync(markdownFile)) {
        setTimeout(() => {
            vscode.commands.executeCommand('vscode.open', vscode.Uri.file(markdownFile));
        }, 1000);
    }
    
    // Auto-show report if enabled
    if (showReport && fs.existsSync(reportFile)) {
        setTimeout(() => {
            vscode.commands.executeCommand('vscode.open', vscode.Uri.file(reportFile));
        }, 2000);
    }
}

export function deactivate() {
    if (outputChannel) {
        outputChannel.dispose();
    }
    console.log('Docx2MD Converter extension deactivated');
}
