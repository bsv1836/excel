# React Excel Clone

A high-performance, browser-based Microsoft Excel alternative built entirely as a single-page React application.

This interface accurately replicates the classic Microsoft Office environment with a fully functional ribbon, formula calculations, tab management, and native `.xlsx` and `.csv` compatibility.

## 🚀 Features

* **Authentic Ribbon Interface**: Features the recognizable Home tab, providing instant access to font manipulation and alignment tools.
* **Typographic Control**: Format individual cells with dynamic font-family and font-size selectors ranging up to 72 points, alongside classic Bold, Italic, and Underline formatting.
* **Full File I/O**: Interoperable with actual Excel files. Import your `.xlsx` or `.csv` sheets instantly and export your edits back out cleanly without data loss.
* **Professional Formula Engine**: Integrates `HyperFormula` internally, offering robust processing of math references, ranges, and functions underneath the hood.
* **Search Context**: Quickly find data with the built-in sheet scanner mapped to a dynamic cell-jumper.
* **Print-Ready PDFs**: Tailored `@media print` rules ensure your spreadsheets break out of screen-constraints when printing directly to a PDF format.

## 🛠 Tech Stack

* **Framework**: React 18 / Vite
* **Styling**: TailwindCSS Custom Configuration
* **Core Logic Engine**: HyperFormula (Spreadsheet Data/Formulas)
* **File Parser**: SheetJS (XLSX, XLS, CSV parsing)
* **Icons**: Lucide-React

## 📦 Local Development

1. **Clone the repository:**
   ```bash
   git clone https://github.com/bsv1836/excel.git
   ```
2. **Install dependencies:**
   ```bash
   npm install
   ```
3. **Start the local server:**
   ```bash
   npm run dev
   ```

## 🌐 Live Deployment
This web application is completely client-side. There is no server/database backend required. It relies entirely on browser storage, making it lightning fast and entirely safe for deployment on any static platform. All files are rendered purely in your browser.
