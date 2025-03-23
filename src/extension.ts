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
        console.log('Creating ODSDocument for:', odsPath);
        this.originalPath = odsPath;
        this.tempDir = path.join(os.tmpdir(), `vscode-ods-${path.basename(odsPath)}-${Date.now()}`);
        console.log('Temp directory:', this.tempDir);
    }

    async open() {
        try {
            console.log('Opening ODS document');
            
            // Verify file exists
            if (!fs.existsSync(this.originalPath)) {
                throw new Error(`File not found: ${this.originalPath}`);
            }

            // Create temp directory
            console.log('Creating temp directory');
            fs.mkdirSync(this.tempDir, { recursive: true });

            // Unzip ODS file
            console.log('Unzipping ODS file');
            const zip = new AdmZip(this.originalPath);
            zip.extractAllTo(this.tempDir, true);

            // Read content.xml
            console.log('Reading content.xml');
            const contentPath = path.join(this.tempDir, 'content.xml');
            if (!fs.existsSync(contentPath)) {
                throw new Error('content.xml not found in ODS file');
            }
            const xmlContent = fs.readFileSync(contentPath, 'utf-8');

            // Show spreadsheet view
            console.log('Creating spreadsheet view');
            SpreadsheetPanel.createOrShow(path.dirname(this.originalPath), xmlContent);

            // Register save handler
            console.log('Registering save handler');
            this.disposables.push(
                vscode.commands.registerCommand('vscode-ods-editor.saveChanges', async (updatedXml: string) => {
                    await this.saveChanges(updatedXml);
                })
            );
        } catch (error) {
            console.error('Failed to open ODS file:', error);
            vscode.window.showErrorMessage(`Failed to open ODS file: ${error}`);
            throw error; // Re-throw to handle in caller
        }
    }

    private async saveChanges(updatedXml: string) {
        try {
            console.log('Saving changes to ODS document');
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
            console.error('Failed to save changes:', error);
            vscode.window.showErrorMessage(`Failed to save changes: ${error}`);
        }
    }

    dispose() {
        console.log('Disposing of ODS document');
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
    console.log('ODS Editor extension is now active');

    // Register command handler
    let disposable = vscode.commands.registerCommand('vscode-ods-editor.openODS', async (uri?: vscode.Uri) => {
        try {
            console.log('Command executed: vscode-ods-editor.openODS');
            console.log('URI:', uri?.fsPath);
            
            let filePath: string;
            
            if (uri) {
                // Opened from context menu
                filePath = uri.fsPath;
                console.log('Opening from context menu:', filePath);
            } else {
                // Opened from command palette
                console.log('Opening file dialog');
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
                    console.log('No file selected');
                    return;
                }

                filePath = files[0].fsPath;
                console.log('Selected file:', filePath);
            }

            // Create and open document
            console.log('Creating ODSDocument');
            const doc = new ODSDocument(filePath);
            console.log('Opening document');
            await doc.open();

            // Register for cleanup
            context.subscriptions.push({
                dispose: () => doc.dispose()
            });

        } catch (error) {
            console.error('Failed to open ODS file:', error);
            vscode.window.showErrorMessage(`Failed to open ODS file: ${error}`);
        }
    });

    // Register file handler for .ods files
    context.subscriptions.push(
        vscode.workspace.registerTextDocumentContentProvider('ods', {
            provideTextDocumentContent(uri: vscode.Uri): string {
                return ''; // We handle the content through our custom editor
            }
        })
    );

    // Add command to the context menu for .ods files
    context.subscriptions.push(
        vscode.commands.registerCommand('vscode-ods-editor.openFromExplorer', async (uri: vscode.Uri) => {
            console.log('Command executed: vscode-ods-editor.openFromExplorer');
            console.log('URI:', uri.fsPath);
            const doc = new ODSDocument(uri.fsPath);
            await doc.open();
            context.subscriptions.push({
                dispose: () => doc.dispose()
            });
        })
    );

    context.subscriptions.push(disposable);
    console.log('ODS Editor extension initialized successfully');
}

export function deactivate() {}