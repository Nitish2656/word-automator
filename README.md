# Word Automator Desktop App

![Electron](https://img.shields.io/badge/Electron-47848F?style=for-the-badge&logo=electron&logoColor=white) 
![NodeJS](https://img.shields.io/badge/Node.js-339933?style=for-the-badge&logo=nodedotjs&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)

A premium, high-performance desktop application built with **Electron.js** and **Node.js**...

## 🚀 Features

- **Multi-Template Support**: Seamlessly switch between MBP-1 and DIR-8 templates using a smooth, tabbed interface.
- **Dual Data Entry**: 
  - **Excel Automation**: Bulk upload data using `.xlsx` files.
  - **Manual Entry**: Interactive forms for quick, single-document generation.
- **Intelligent Template Parsing**: 
  - Advanced XML pre-processing to handle complex Word placeholders (like duplicate `(*)` markers).
  - Preserves formatting, tables, and styles from the original Word templates.
- **Modern UI/UX**: Dark-themed, glassmorphic design with interactive states and smooth transitions.
- **Portable Setup**: Generates a single-file `Setup.exe` for easy distribution.

## 🛠️ Technology Stack

- **Core**: Electron.js, Node.js
- **Templating**: [docxtemplater](https://docxtemplater.com/), PizZip
- **Data Parsing**: [SheetJS (XLSX)](https://sheetjs.com/)
- **Styling**: Vanilla CSS (Premium Dark Theme)

## 📦 How to Use

1. **Select Template**: Set the path to your `.docx` template files.
2. **Input Data**: Either drag & drop an Excel file or fill out the manual form.
3. **Generate**: Hit the "Generate Word Documents" button. The files will be saved in your specified output directory.

## 📚 Documentation

For a deep dive into the application's architecture, template logic, and developer guides, please refer to the:
- **[In-Depth Technical Documentation (DOCUMENTATION.md)](./DOCUMENTATION.md)**

## 🔧 Installation for Developers

```bash
# Clone the repository
git clone https://github.com/Nitish2656/word-automator.git

# Install dependencies
npm install

# Run the app
npm start

# Build the installer
npm run build
```

---
