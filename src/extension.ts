import * as vscode from 'vscode';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import * as AdmZip from 'adm-zip';
import { SpreadsheetPanel } from './spreadsheetPanel';

class ODSDocument {
    private originalPath: string;
    private tempDir: string;
    private disposables: vscode.Disposable[] = [];

    constructor(odsPath: string) {
        this.originalPath = odsPath;
        this.tempDir = path.join(os.tmpdir(), `vscode-ods-${path.basename(odsPath)}-${Date.now()}`);
    }

    async open() {
        try {
            // Create temp directory
            fs.mkdirSync(this.tempDir, { recursive: true });

            // Unzip ODS file
            const zip = new AdmZip(this.originalPath);
            zip.extractAllTo(this.tempDir, true);

            // Read content.xml
            const contentPath = path.join(this.tempDir, 'content.xml');
            const xmlContent = fs.readFileSync(contentPath, 'utf-8');

            // Show spreadsheet view
            SpreadsheetPanel.createOrShow(path.dirname(this.originalPath), xmlContent);

            // Register save handler
            this.disposables.push(
                vscode.commands.registerCommand('vscode-ods-editor.saveChanges', async (updatedXml: string) => {
                    await this.saveChanges(updatedXml);
                })
            );
        } catch (error) {
            vscode.window.showErrorMessage(`Failed to open ODS file: ${error}`);
        }
    }

    private async saveChanges(updatedXml: string) {
        try {
            // Write updated content.xml
            const contentPath = path.join(this.tempDir, 'content.xml');
            fs.writeFileSync(contentPath, updatedXml, 'utf-8');

            // Create new zip with updated content
            const zip = new AdmZip();
            const files = fs.readdirSync(this.tempDir);
            
            for (const file of files) {
                const filePath = path.join(this.tempDir, file);
                zip.addLocalFile(filePath);
            }

            // Save the updated ODS file
            zip.writeZip(this.originalPath);

            // Open in LibreOffice
            const terminal = vscode.window.createTerminal('Open in LibreOffice');
            if (process.platform === 'win32') {
                terminal.sendText(`start "" "${this.originalPath}"`);
            } else if (process.platform === 'darwin') {
                terminal.sendText(`open "${this.originalPath}"`);
            } else {
                terminal.sendText(`xdg-open "${this.originalPath}"`);
            }
            terminal.dispose();

            vscode.window.showInformationMessage('Spreadsheet saved successfully!');
        } catch (error) {
            vscode.window.showErrorMessage(`Failed to save changes: ${error}`);
        }
    }

    dispose() {
        // Clean up temp directory
        try {
            if (fs.existsSync(this.tempDir)) {
                fs.rmSync(this.tempDir, { recursive: true });
            }
        } catch (error) {
            console.error('Failed to clean up temp directory:', error);
        }

        // Dispose of all disposables
        this.disposables.forEach(d => d.dispose());
    }
}

export function activate(context: vscode.ExtensionContext) {
    // Register command handler
    let disposable = vscode.commands.registerCommand('vscode-ods-editor.openODS', async (uri?: vscode.Uri) => {
        try {
            let filePath: string;
            
            if (uri) {
                // Opened from context menu
                filePath = uri.fsPath;
            } else {
                // Opened from command palette
                const files = await vscode.window.showOpenDialog({
                    canSelectFiles: true,
                    canSelectFolders: false,
                    canSelectMany: false,
                    filters: {
                        'ODS Files': ['ods']
                    },
                    title: 'Select ODS File to Open'
                });

                if (!files || files.length === 0) {
                    return;
                }

                filePath = files[0].fsPath;
            }

            // Create and open document
            const doc = new ODSDocument(filePath);
            await doc.open();

            // Register for cleanup
            context.subscriptions.push({
                dispose: () => doc.dispose()
            });

        } catch (error) {
            vscode.window.showErrorMessage(`Failed to open ODS file: ${error}`);
        }
    });

    context.subscriptions.push(disposable);
}

export function deactivate() {}