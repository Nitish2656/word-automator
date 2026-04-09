# Word Automator Desktop App

A premium, high-performance desktop application built with Electron.js designed to automate the generation of Word documents. specifically tailored for **MBP-1** and **DIR-8** forms.

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
