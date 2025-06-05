// The module 'vscode' contains the VS Code extensibility API
// Import the module and reference it with the alias vscode in your code below
import * as vscode from 'vscode';
import { execFile } from 'child_process';
import * as path from 'path';

// This method is called when your extension is activated
// Your extension is activated the very first time the command is executed
export function activate(context: vscode.ExtensionContext) {

	// Use the console to output diagnostic information (console.log) and errors (console.error)
	// This line of code will only be executed once when your extension is activated
	console.log('Congratulations, your extension "docx2mdconverter" is now active!');

	// The command has been defined in the package.json file
	// Now provide the implementation of the command with registerCommand
	// The commandId parameter must match the command field in package.json
	const disposable = vscode.commands.registerCommand('docx2mdconverter.helloWorld', () => {
		// The code you place here will be executed every time your command is executed
		// Display a message box to the user
		vscode.window.showInformationMessage('Hello World from Docx2MDConverter!');
	});

	context.subscriptions.push(disposable);

	let convertDisposable = vscode.commands.registerCommand('docx2mdconverter.convertDocxToMarkdown', async () => {
		// Prompt user to select a DOCX file
		const docxUris = await vscode.window.showOpenDialog({
			canSelectMany: false,
			openLabel: 'Select DOCX file',
			filters: { 'Word Documents': ['docx'] }
		});
		if (!docxUris || docxUris.length === 0) {
			vscode.window.showWarningMessage('No DOCX file selected.');
			return;
		}
		const docxPath = docxUris[0].fsPath;

		// Prompt user to select output directory
		const outputUris = await vscode.window.showOpenDialog({
			canSelectFolders: true,
			canSelectFiles: false,
			canSelectMany: false,
			openLabel: 'Select Output Directory'
		});
		if (!outputUris || outputUris.length === 0) {
			vscode.window.showWarningMessage('No output directory selected.');
			return;
		}
		const outputDir = outputUris[0].fsPath;

		// Find the Python script in the workspace
		const pythonScript = path.join(context.extensionPath, 'docx_to_markdown_converter.py');

		// Output channel for results
		const outputChannel = vscode.window.createOutputChannel('Docx2MD');
		outputChannel.show(true);
		outputChannel.appendLine(`Converting: ${docxPath}`);

		// Run the Python script
		execFile('python', [pythonScript, docxPath, '--output', outputDir], (error, stdout, stderr) => {
			if (error) {
				outputChannel.appendLine(`Error: ${error.message}`);
				vscode.window.showErrorMessage('Conversion failed. See output for details.');
				return;
			}
			if (stderr) {
				outputChannel.appendLine(`stderr: ${stderr}`);
			}
			outputChannel.appendLine(stdout);
			vscode.window.showInformationMessage('DOCX to Markdown conversion complete!');
		});
	});

	context.subscriptions.push(convertDisposable);
}

// This method is called when your extension is deactivated
export function deactivate() {}
