let documentData = [];
let activeTemplateType = 'mbp-1'; // mbp-1 or dir-8

// DOM Elements
const templatePathInput = document.getElementById('template-path');
const btnSelectTemplate = document.getElementById('btn-select-template');
const outputPathInput = document.getElementById('output-path');
const btnSelectOutput = document.getElementById('btn-select-output');

const excelDropZone = document.getElementById('excel-drop-zone');
const excelFileInput = document.getElementById('excel-file-input');
const btnBrowseExcel = document.getElementById('btn-browse-excel');
const excelStatus = document.getElementById('excel-status');

const tabBtns = document.querySelectorAll('.tab-btn');
const manualDataForms = document.querySelectorAll('.manual-data-form');
const btnAddManuals = document.querySelectorAll('.btn-add-manual');

const dataTable = document.getElementById('data-table');
const tableHeadRow = document.getElementById('table-head-row');
const tableBody = document.getElementById('table-body');
const dataCount = document.getElementById('data-count');
const btnClearData = document.getElementById('btn-clear-data');

const btnGenerate = document.getElementById('btn-generate');
const processResult = document.getElementById('process-result');

// Set Defaults
let currentTemplatePaths = {
    'mbp-1': "Z:\\Documents\\Blanked\\Blank_MBP-1_.docx",
    'dir-8': "Z:\\Documents\\Blanked\\Blank_DIR-8.docx"
};
let currentOutputPath = "Z:\\Documents\\Filled";
templatePathInput.value = currentTemplatePaths[activeTemplateType];
outputPathInput.value = currentOutputPath;

// Setup IPC Calls
btnSelectTemplate.addEventListener('click', async () => {
    const path = await window.electronAPI.selectFile({
        filters: [{ name: 'Word Documents', extensions: ['docx'] }]
    });
    if (path) {
        currentTemplatePaths[activeTemplateType] = path;
        templatePathInput.value = path;
    }
});

btnSelectOutput.addEventListener('click', async () => {
    const path = await window.electronAPI.selectDirectory();
    if (path) {
        currentOutputPath = path;
        outputPathInput.value = path;
    }
});

// Tab Switching
tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
        tabBtns.forEach(b => b.classList.remove('active'));
        manualDataForms.forEach(f => f.style.display = 'none');
        
        btn.classList.add('active');
        activeTemplateType = btn.dataset.tab;
        
        document.getElementById(`form-${activeTemplateType}`).style.display = 'block';
        templatePathInput.value = currentTemplatePaths[activeTemplateType];
    });
});

// Excel Upload Logic
btnBrowseExcel.addEventListener('click', () => excelFileInput.click());

excelFileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleExcelFile(e.target.files[0]);
    }
});

excelDropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    excelDropZone.classList.add('drag-active');
});

excelDropZone.addEventListener('dragleave', () => {
    excelDropZone.classList.remove('drag-active');
});

excelDropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    excelDropZone.classList.remove('drag-active');
    if (e.dataTransfer.files.length > 0) {
        handleExcelFile(e.dataTransfer.files[0]);
    }
});

function handleExcelFile(file) {
    excelStatus.className = 'status-msg';
    excelStatus.textContent = `Reading ${file.name}...`;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const json = XLSX.utils.sheet_to_json(worksheet);

            if (json.length > 0) {
                // Attach the currently active template type to these rows
                const rowsWithType = json.map(row => ({ ...row, _templateType: activeTemplateType }));
                documentData = [...documentData, ...rowsWithType];
                updateTable();
                excelStatus.textContent = `Added ${json.length} rows as ${activeTemplateType}.`;
                excelStatus.className = 'status-msg success';
            } else {
                excelStatus.textContent = 'Excel file is empty.';
                excelStatus.className = 'status-msg error';
            }
        } catch (error) {
            excelStatus.textContent = `Error reading Excel: ${error.message}`;
            excelStatus.className = 'status-msg error';
        }
    };
    reader.readAsArrayBuffer(file);
}

// Manual Entry Logic
btnAddManuals.forEach(btn => {
    btn.addEventListener('click', () => {
        const formId = btn.dataset.target;
        const form = document.getElementById(formId);
        const formData = new FormData(form);
        const rowData = { _templateType: activeTemplateType };
        let hasData = false;

        for (let [key, value] of formData.entries()) {
            if (value.trim() !== '') {
                rowData[key] = value.trim();
                hasData = true;
            }
        }

        const requiredKey = activeTemplateType === 'mbp-1' ? 'name of the Company' : 'Name of Company';

        if (hasData && rowData[requiredKey]) {
            documentData.push(rowData);
            updateTable();
            form.reset();
        } else {
            alert("Please enter at least the Company Name.");
        }
    });
});

// Table & UI Updates
function updateTable() {
    dataCount.textContent = documentData.length;
    tableBody.innerHTML = '';

    if (documentData.length > 0) {
        btnGenerate.disabled = false;
        
        // Dynamically update headers based on all unique keys
        const allKeys = new Set();
        documentData.forEach(row => Object.keys(row).forEach(k => {
            if (k !== '_templateType') allKeys.add(k);
        }));
        
        tableHeadRow.innerHTML = '<th>Type</th>';
        allKeys.forEach(key => {
            const th = document.createElement('th');
            th.textContent = key;
            tableHeadRow.appendChild(th);
        });

        const thAction = document.createElement('th');
        thAction.textContent = 'Action';
        tableHeadRow.appendChild(thAction);

        documentData.forEach((row, index) => {
            const tr = document.createElement('tr');
            
            const tdType = document.createElement('td');
            tdType.textContent = row._templateType.toUpperCase();
            tr.appendChild(tdType);

            allKeys.forEach(key => {
                const td = document.createElement('td');
                td.textContent = row[key] || '';
                tr.appendChild(td);
            });

            const tdAction = document.createElement('td');
            const btnDelete = document.createElement('button');
            btnDelete.textContent = 'X';
            btnDelete.className = 'btn btn-outline-danger btn-sm';
            btnDelete.onclick = () => {
                documentData.splice(index, 1);
                updateTable();
            };
            tdAction.appendChild(btnDelete);
            tr.appendChild(tdAction);

            tableBody.appendChild(tr);
        });
    } else {
        btnGenerate.disabled = true;
        tableHeadRow.innerHTML = '<th>Company</th><th>Director</th><th>DIN</th>';
    }
}

btnClearData.addEventListener('click', () => {
    documentData = [];
    updateTable();
    excelStatus.textContent = '';
    processResult.textContent = '';
});

// Generate Documents
btnGenerate.addEventListener('click', async () => {
    if (documentData.length === 0) return;

    btnGenerate.disabled = true;
    processResult.textContent = 'Generating documents...';
    processResult.className = 'status-msg';

    // Group data by template so main process can run efficiently
    const result = await window.electronAPI.generateDocs({
        data: documentData,
        outputDir: currentOutputPath,
        templatePaths: currentTemplatePaths
    });

    if (result.success) {
        processResult.textContent = `Successfully generated ${result.files.length} document(s) in ${currentOutputPath}`;
        processResult.className = 'status-msg success';
        documentData = [];
        updateTable();
    } else {
        processResult.textContent = `Error: ${result.error}`;
        processResult.className = 'status-msg error';
        btnGenerate.disabled = false;
    }
});
