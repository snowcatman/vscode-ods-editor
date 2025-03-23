const fs = require('fs');
const path = require('path');
const os = require('os');

function cleanupTempFiles() {
    const tempDir = os.tmpdir();
    const files = fs.readdirSync(tempDir);
    
    // Find and remove all vscode-ods-* temporary directories
    files.forEach(file => {
        if (file.startsWith('vscode-ods-')) {
            const fullPath = path.join(tempDir, file);
            try {
                fs.rmSync(fullPath, { recursive: true, force: true });
                console.log(`Removed temporary directory: ${fullPath}`);
            } catch (error) {
                console.error(`Failed to remove ${fullPath}: ${error}`);
            }
        }
    });
}

// Clean up temp files
cleanupTempFiles();

// Log uninstall completion
console.log('ODS Editor extension uninstalled successfully'); 