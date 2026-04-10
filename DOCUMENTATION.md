# 📘 In-Depth Project Documentation: Word Automator

Welcome to the comprehensive technical documentation for the **Word Automator** desktop application. This guide covers everything from architecture and logic to troubleshooting and deployment.

---

## 📑 Table of Contents
1. [Architecture Overview](#architecture-overview)
2. [Core features Logic](#core-features-logic)
   - [Excel Automation](#excel-automation)
   - [Manual Entry](#manual-entry)
3. [Template Engine & Custom Processing](#template-engine--custom-processing)
   - [MBP-1 Logic](#mbp-1-logic)
   - [DIR-8 Logic (*) Sequential Mapping](#dir-8-logic--sequential-mapping)
4. [File Structure](#file-structure)
5. [Developer Guide](#developer-guide)
   - [Installation](#installation)
   - [Building the App](#building-the-app)
6. [Troubleshooting & FAQs](#troubleshooting--faqs)

---

## 🏗 Architecture Overview

Word Automator is built using the **Electron.js** framework, which allows build cross-platform desktop apps with web technologies.

- **Main Process (`main.js`)**: Acts as the backend. It has direct access to the Node.js filesystem API. It handles the actual Word document generation and sensitive file operations.
- **Renderer Process (`renderer.js`)**: Acts as the frontend. It runs inside the Chrome window, handles user interactions, parses Excel files, and sends data to the main process via IPC.
- **Preload Script (`preload.js`)**: A secure bridge that exposes only specific, safe functions from the Main process to the Renderer process using `contextBridge`.

---

## ⚙️ Core Features Logic

### Excel Automation
Using the `xlsx` (SheetJS) library, the app reads uploaded `.xlsx` files and converts them into JSON objects. 
- **Mapping**: The app looks for headers in your Excel file that match the variable names in the Word template (e.g., "name of the Company").
- **Bulk Processing**: Multiple rows in an Excel file are processed as a queue. Each row results in a separate generated Word document.

### Manual Entry
The app provides two interactive forms (MBP-1 and DIR-8). 
- **Validation**: Ensures a "Company Name" is provided before adding to the queue.
- **Queue System**: Users can add multiple manual entries to a "Data Table" before hitting generate, allowing for efficient bulk production even without an Excel file.

---

## 📝 Template Engine & Custom Processing

The app uses `docxtemplater` with the `PizZip` library. We utilize custom **delimiters** `(` and `)` to match your existing Word formats.

### MBP-1 Logic
In the MBP-1 template, variables like `(name of the Company)` are replaced directly. Special handling is added in `main.js` to ensure that standard parentheses used in text (like section numbers) are not accidentally deleted by the engine.

### DIR-8 Logic (*) Sequential Mapping
The DIR-8 template uses generic `(*)` symbols for many different fields. Standard templating cannot distinguish between them. 
**Our Unique Solution**:
- Before the templating engine runs, the app performs a **Raw XML Pre-process**.
- It opens the Word document's internal XML and sequentially replaces the specific `<w:t>(*)</w:t>` tags in a fixed order:
  1. Registration No.
  2. Nominal Capital
  3. Paid-up Capital
  4. Name of Company
  5. Registered Office Address
  6. Date
  7. Place

---

## 📂 File Structure

```text
word-automator/
├── main.js           # Main backend/engine
├── renderer.js       # Frontend UI logic
├── index.html        # UI Structure
├── style.css         # UI Design
├── preload.js        # Secure IPC bridge
├── package.json      # Dependencies & Build config
├── templates/        # Default .docx template files
│   ├── Blank_MBP-1_.docx
│   └── Blank_DIR-8.docx
└── dist/             # Generated executables (after build)
```

---

## 🛠 Developer Guide

### Installation
Ensure you have Node.js installed, then run:
```bash
npm install
```

### Building the App
To create a standalone `.exe` for Windows:
```bash
npm run build
```
The output will be in the `dist/word-automator-win32-x64` folder.

---

## ❓ Troubleshooting & FAQs

**Q: Why does my DIR-8 file look empty?**
A: Ensure your input file is a `.docx`. If you have an old `.doc` file, you must save it as "Word Document (.docx)" first.

**Q: Can I change the variable names?**
A: Yes. If you change a variable in the Word template (e.g., from `(Date)` to `(Current Date)`), simply update the `name` attribute of the corresponding input field in `index.html`.

**Q: The generation is failing on certain files.**
A: Check if the template file is currently open in Microsoft Word. Sometimes Word "locks" the file, preventing the app from reading it.

---
*Documentation curated by Antigravity for Nitish.*
