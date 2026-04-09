const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1000,
        height: 700,
        webPreferences: {
            preload: path.join(__dirname, 'preload.js'),
            nodeIntegration: false,
            contextIsolation: true
        }
    });

    mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
    createWindow();

    app.on('activate', function () {
        if (BrowserWindow.getAllWindows().length === 0) createWindow();
    });
});

app.on('window-all-closed', function () {
    if (process.platform !== 'darwin') app.quit();
});

// IPC Handler to select a directory for saving files
ipcMain.handle('select-directory', async () => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openDirectory']
    });
    return result.filePaths[0]; // will be undefined if user cancelled
});

// IPC Handler to select the template file (if not using default)
ipcMain.handle('select-file', async (event, options) => {
    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile'],
        filters: options?.filters || []
    });
    return result.filePaths[0];
});

// IPC Handler to generate word documents
ipcMain.handle('generate-docs', async (event, { data, outputDir, templatePaths }) => {
    try {
        const generatedFiles = [];

        // Group data by template type
        const groupedData = data.reduce((acc, row) => {
            const type = row._templateType || 'mbp-1';
            if (!acc[type]) acc[type] = [];
            acc[type].push(row);
            return acc;
        }, {});

        for (const type in groupedData) {
            const templatePath = templatePaths[type];
            if (!fs.existsSync(templatePath)) {
                throw new Error(`Template file not found at ${templatePath}`);
            }

            const content = fs.readFileSync(templatePath, 'binary');
            const rows = groupedData[type];

            rows.forEach((row, index) => {
                const zip = new PizZip(content);
                let docXml = zip.file("word/document.xml").asText();

                // Template specific pre-processing
                if (type === 'dir-8') {
                    // In DIR-8, the (*) markers are exactly in their own <w:t>(*)</w:t> tags in a specific order:
                    // 1: Registration, 2: Nominal, 3: Paid-up, 4: Company, 5: Address, 6: Date, 7: Place
                    const r = (val) => val ? `<w:t>${val}</w:t>` : `<w:t>(*)</w:t>`;
                    
                    docXml = docXml.replace('<w:t>(*)</w:t>', r(row['Registration No. of Company']));
                    docXml = docXml.replace('<w:t>(*)</w:t>', r(row['Nominal Capital']));
                    docXml = docXml.replace('<w:t>(*)</w:t>', r(row['Paid-up Capital']));
                    docXml = docXml.replace('<w:t>(*)</w:t>', r(row['Name of Company']));
                    docXml = docXml.replace('<w:t>(*)</w:t>', r(row['Address of its Registered Office']));
                    docXml = docXml.replace('<w:t>(*)</w:t>', r(row['Date']));
                    docXml = docXml.replace('<w:t>(*)</w:t>', r(row['Place']));
                } else {
                    // Common pre-processing for shared (*) tags in MBP-1
                    if (row['Date']) docXml = docXml.replace(/Date:\s*\(\*\)/g, `Date: ${row['Date']}`);
                    if (row['Place']) docXml = docXml.replace(/Place:\s*\(\*\)/g, `Place: ${row['Place']}`);
                }

                zip.file("word/document.xml", docXml);

                const doc = new Docxtemplater(zip, {
                    paragraphLoop: true,
                    linebreaks: true,
                    delimiters: { start: '(', end: ')' },
                    nullGetter: function(part) {
                        if (!part.module) {
                            return "(" + part.value + ")";
                        }
                        if (part.module === "rawxml") {
                            return "";
                        }
                        return "";
                    }
                });

                // Map alternate tags for MBP-1 and DIR-8 variations in the Word Document
                if (row['name of the Director']) row['Name of Director'] = row['name of the Director'];
                
                // For DIR-8, the form uses 'Name of Company', but template has '(Name of the Company)'
                if (row['Name of Company']) row['name of the Company'] = row['Name of Company'];
                if (row['name of the Company']) row['Name of Company'] = row['name of the Company'];

                // Handle Father name casing
                if (row['name of the Father of Director']) {
                    row['Name of Father of Director'] = row['name of the Father of Director'];
                }
                
                // Handle DIN variations
                if (row['DIN no.']) {
                    row['DIN No. of Director'] = row['DIN no.'];
                    row['DIN of Director'] = row['DIN no.'];
                }

                doc.render(row);

                const buf = doc.getZip().generate({
                    type: 'nodebuffer',
                    compression: 'DEFLATE',
                });

                const companyName = row['name of the Company'] || row['Name of Company'] || row.Company || `Doc_${index + 1}`;
                const fileName = `${companyName.replace(/[^a-z0-9]/gi, '_').toLowerCase()}_${type.toUpperCase()}.docx`;
                const outputPath = path.join(outputDir, fileName);

                fs.writeFileSync(outputPath, buf);
                generatedFiles.push(outputPath);
            });
        }

        return { success: true, files: generatedFiles };
    } catch (error) {
        console.error('Error generating document:', error);
        return { success: false, error: error.message };
    }
});
