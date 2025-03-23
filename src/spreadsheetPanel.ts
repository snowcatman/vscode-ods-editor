import * as vscode from 'vscode';
import * as path from 'path';
import * as fs from 'fs';
import { XMLParser, XMLBuilder } from 'fast-xml-parser';

export class SpreadsheetPanel {
    public static currentPanel: SpreadsheetPanel | undefined;
    private readonly _panel: vscode.WebviewPanel;
    private readonly _extensionPath: string;
    private _disposables: vscode.Disposable[] = [];
    private _currentXMLContent: string = '';

    private constructor(panel: vscode.WebviewPanel, extensionPath: string) {
        this._panel = panel;
        this._extensionPath = extensionPath;

        this._panel.onDidDispose(() => this.dispose(), null, this._disposables);

        // Handle messages from the webview
        this._panel.webview.onDidReceiveMessage(
            async message => {
                switch (message.command) {
                    case 'cellEdit':
                        await this._handleCellEdit(message.cell, message.value);
                        break;
                    case 'ready':
                        // Send initial content when webview is ready
                        this._panel.webview.postMessage({
                            command: 'update',
                            content: this._parseODSContent(this._currentXMLContent)
                        });
                        break;
                }
            },
            null,
            this._disposables
        );
    }

    private _parseODSContent(xmlContent: string) {
        try {
            const parser = new XMLParser({
                ignoreAttributes: false,  
                attributeNamePrefix: '@_',
                parseAttributeValue: true,
                parseTagValue: false,
                trimValues: true,
                processEntities: true,
                htmlEntities: true
            });
            
            const xmlObj = parser.parse(xmlContent);
            console.log('Raw XML:', xmlObj);

            const spreadsheet = xmlObj['office:document-content']
                ['office:body']
                ['office:spreadsheet'];

            const tables = Array.isArray(spreadsheet['table:table']) 
                ? spreadsheet['table:table'] 
                : [spreadsheet['table:table']];

            // Create unique sheet names if none exist
            const sheets = tables.map((table, index) => {
                const sheetName = table['@_table:name'] || `Sheet${index + 1}`;
                return this._convertTable(table, sheetName);
            });

            console.log('Sheets:', sheets);
            return { sheets };
        } catch (error) {
            console.error('Error parsing ODS content:', error);
            vscode.window.showErrorMessage(`Failed to parse spreadsheet: ${error}`);
            return { sheets: [{ name: 'Sheet1', rows: [[{ value: 'Error loading spreadsheet' }]] }] };
        }
    }

    private _convertTable(table: any, sheetName: string) {
        try {
            const rows = Array.isArray(table['table:table-row']) 
                ? table['table:table-row'] 
                : [table['table:table-row']];

            return {
                name: sheetName,
                rows: rows.map(this._convertRow.bind(this))
            };
        } catch (error) {
            console.error('Error converting table:', error);
            return { name: sheetName, rows: [[{ value: 'Error converting table' }]] };
        }
    }

    private _convertRow(row: any) {
        try {
            if (!row['table:table-cell']) {
                return [{ value: '' }];
            }

            const cells = Array.isArray(row['table:table-cell']) 
                ? row['table:table-cell'] 
                : [row['table:table-cell']];

            // Handle repeated columns
            const expandedCells = [];
            for (const cell of cells) {
                const repeatCount = parseInt(cell['@_table:number-columns-repeated'] || '1', 10);
                for (let i = 0; i < repeatCount; i++) {
                    expandedCells.push(this._convertCell(cell));
                }
            }

            return expandedCells;
        } catch (error) {
            console.error('Error converting row:', error);
            return [{ value: 'Error converting row' }];
        }
    }

    private _convertCell(cell: any) {
        try {
            let value = '';
            
            // Get cell value from text:p if it exists
            if (cell['text:p']) {
                value = Array.isArray(cell['text:p']) 
                    ? cell['text:p'].join(' ') 
                    : cell['text:p'];
            }

            return {
                value: value,
                type: cell['@_office:value-type'] || 'string',
                formula: cell['@_table:formula'] || '',
                style: cell['@_table:style-name'] || ''
            };
        } catch (error) {
            console.error('Error converting cell:', error);
            return { value: 'Error', type: 'string', formula: '', style: '' };
        }
    }

    private async _handleCellEdit(cellId: string, value: string) {
        try {
            const [sheet, row, col] = cellId.split(':').map(Number);
            const parser = new XMLParser({
                ignoreAttributes: false,
                attributeNamePrefix: '',
                parseAttributeValue: true,
                parseTagValue: true,
                trimValues: true
            });
            
            const xmlObj = parser.parse(this._currentXMLContent);
            const spreadsheet = xmlObj['office:document-content']['office:body']['office:spreadsheet'];
            const tables = spreadsheet['table:table'];
            
            // Update cell in the XML structure
            const targetTable = Array.isArray(tables) ? tables[sheet] : tables;
            const rows = targetTable['table:table-row'];
            const targetRow = Array.isArray(rows) ? rows[row] : rows;
            const cells = targetRow['table:table-cell'];
            
            if (Array.isArray(cells)) {
                cells[col] = {
                    ...cells[col],
                    'office:value-type': 'string',
                    'text:p': value
                };
            } else if (col === 0) {
                targetRow['table:table-cell'] = {
                    ...cells,
                    'office:value-type': 'string',
                    'text:p': value
                };
            }
            
            // Build updated XML
            const builder = new XMLBuilder({
                ignoreAttributes: false,
                attributeNamePrefix: '',
                format: true
            });
            const updatedXml = builder.build(xmlObj);
            
            // Save changes
            this._currentXMLContent = updatedXml;
            
            // Notify extension to save changes
            vscode.commands.executeCommand('vscode-ods-editor.saveChanges', updatedXml);
        } catch (error) {
            console.error('Error handling cell edit:', error);
            vscode.window.showErrorMessage(`Failed to update cell: ${error}`);
        }
    }

    public static createOrShow(extensionPath: string, xmlContent: string) {
        const column = vscode.window.activeTextEditor
            ? vscode.window.activeTextEditor.viewColumn
            : undefined;

        if (SpreadsheetPanel.currentPanel) {
            SpreadsheetPanel.currentPanel._panel.reveal(column);
            SpreadsheetPanel.currentPanel.updateContent(xmlContent);
            return;
        }

        const panel = vscode.window.createWebviewPanel(
            'odsSpreadsheet',
            'ODS Spreadsheet',
            column || vscode.ViewColumn.One,
            {
                enableScripts: true,
                localResourceRoots: [vscode.Uri.file(path.join(extensionPath, 'media'))]
            }
        );

        SpreadsheetPanel.currentPanel = new SpreadsheetPanel(panel, extensionPath);
        SpreadsheetPanel.currentPanel.updateContent(xmlContent);
    }

    public updateContent(xmlContent: string) {
        this._currentXMLContent = xmlContent;
        this._panel.webview.html = this._getWebviewContent();
    }

    private _getWebviewContent() {
        return `<!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <style>
                body { margin: 0; padding: 10px; font-family: Arial, sans-serif; }
                .spreadsheet { border-collapse: collapse; width: 100%; }
                .spreadsheet td {
                    border: 1px solid #ccc;
                    padding: 4px 8px;
                    min-width: 80px;
                    height: 24px;
                    white-space: nowrap;
                    overflow: hidden;
                    text-overflow: ellipsis;
                }
                .spreadsheet td:focus {
                    outline: 2px solid #0066cc;
                    outline-offset: -1px;
                    white-space: normal;
                    overflow: visible;
                }
                .sheet-tabs {
                    display: flex;
                    margin-top: 10px;
                    border-bottom: 1px solid #ccc;
                    background: #f5f5f5;
                }
                .sheet-tab {
                    padding: 8px 20px;
                    border: 1px solid #ccc;
                    border-bottom: none;
                    margin-right: 2px;
                    cursor: pointer;
                    background: white;
                }
                .sheet-tab.active {
                    background: #e6e6e6;
                    font-weight: bold;
                }
                .spreadsheet tr:first-child td {
                    background: #f5f5f5;
                    font-weight: bold;
                    text-align: center;
                }
                .spreadsheet td:first-child {
                    background: #f5f5f5;
                    font-weight: bold;
                    text-align: center;
                }
            </style>
        </head>
        <body>
            <div id="app">
                <table class="spreadsheet" id="sheet"></table>
                <div class="sheet-tabs" id="tabs"></div>
            </div>
            <script>
                const vscode = acquireVsCodeApi();
                let currentSheet = 0;
                let spreadsheetData = null;

                // Initialize
                document.addEventListener('DOMContentLoaded', () => {
                    vscode.postMessage({ command: 'ready' });
                });

                // Handle messages from extension
                window.addEventListener('message', event => {
                    const message = event.data;
                    switch (message.command) {
                        case 'update':
                            spreadsheetData = message.content;
                            console.log('Received spreadsheet data:', spreadsheetData);
                            renderSpreadsheet();
                            break;
                    }
                });

                function renderSpreadsheet() {
                    if (!spreadsheetData) return;
                    console.log('Rendering spreadsheet:', spreadsheetData);
                    
                    // Render sheet tabs
                    const tabsContainer = document.getElementById('tabs');
                    tabsContainer.innerHTML = spreadsheetData.sheets.map((sheet, idx) => 
                        \`<div class="sheet-tab \${idx === currentSheet ? 'active' : ''}" 
                              onclick="switchSheet(\${idx})">\${sheet.name}</div>\`
                    ).join('');

                    // Render current sheet
                    const sheet = spreadsheetData.sheets[currentSheet];
                    const table = document.getElementById('sheet');
                    
                    // Add column headers (A, B, C, etc.)
                    const maxCols = Math.max(...sheet.rows.map(row => row.length), 26);
                    const headerRow = [''].concat(Array.from({ length: maxCols }, (_, i) => 
                        String.fromCharCode(65 + i)
                    ));
                    
                    table.innerHTML = \`
                        <tr>\${headerRow.map(header => \`<td>\${header}</td>\`).join('')}</tr>
                        \${sheet.rows.map((row, rowIdx) => 
                            \`<tr>
                                <td>\${rowIdx + 1}</td>
                                \${row.map((cell, colIdx) => 
                                    \`<td contenteditable="true" 
                                         data-id="\${currentSheet}:\${rowIdx}:\${colIdx}"
                                         data-type="\${cell.type}"
                                         data-formula="\${cell.formula}">\${cell.value}</td>\`
                                ).join('')}
                            </tr>\`
                        ).join('')}
                    \`;

                    // Add cell edit listeners
                    table.querySelectorAll('td[contenteditable="true"]').forEach(cell => {
                        cell.addEventListener('blur', () => {
                            if (cell.dataset.id) {
                                vscode.postMessage({
                                    command: 'cellEdit',
                                    cell: cell.dataset.id,
                                    value: cell.textContent
                                });
                            }
                        });
                        
                        // Handle Enter key
                        cell.addEventListener('keydown', (e) => {
                            if (e.key === 'Enter' && !e.shiftKey) {
                                e.preventDefault();
                                cell.blur();
                            }
                        });
                    });
                }

                function switchSheet(index) {
                    currentSheet = index;
                    renderSpreadsheet();
                }
            </script>
        </body>
        </html>`;
    }

    public dispose() {
        SpreadsheetPanel.currentPanel = undefined;
        this._panel.dispose();
        while (this._disposables.length) {
            const disposable = this._disposables.pop();
            if (disposable) {
                disposable.dispose();
            }
        }
    }
}