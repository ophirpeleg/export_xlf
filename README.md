# Excel to XLIFF Converter Tool

This tool is designed to streamline translation activities, making it easier to convert between XLIFF and Excel formats. It supports handling multiple XLIFF files, automating feedback workflows, comparing translation files, and creating deployment packages for Tabs and Labels in Salesforce.

## Purpose

The primary purpose of this tool is to support translation workflows by providing easy conversions between XLIFF and Excel formats. It helps in:

- Converting XLIFF files to Excel for easier data handling.
- Converting Excel files back to XLIFF, ready for re-import into translation systems.
- Automating feedback workflows for translations.
- Comparing old and new XLIFF files to identify changes.
- Creating structured deployment packages for Tabs and Labels translations.

## Features & Buttons

### 1. **Excel to XLIFF**
- **Purpose**: Converts an Excel file into an XLIFF file.
- **Usage**: Click to select an Excel file, which will be converted to XLIFF format and saved in your chosen location.

### 2. **XLIFF to Excel**
- **Purpose**: Converts a single XLIFF file into Excel format.
- **Usage**: Click to select an XLIFF file, which will be converted to an Excel workbook. You can save the Excel file to your chosen location.

### 3. **Multiple Files XLIFF to Excel**
- **Purpose**: Converts multiple XLIFF files to separate Excel files.
- **Usage**: Select multiple XLIFF files to process at once. The tool will create individual Excel files for each, saving them in a specified folder organized by target language.

### 4. **Feedback File Automation**
- **Purpose**: Automates the feedback process by generating a report with translation length feedback.
- **Usage**: Select the source language XLIFF file and an optional English XLIFF file (for reference translations). The tool will create an Excel file with columns for feedback on translation length and a field for customer feedback.

### 5. **Create Package (Tabs and Labels)**
- **Purpose**: Creates deployment packages for translated Tabs and Labels.
- **Usage**: Select multiple `.objectTranslation` files to package by language. The tool will remove unnecessary sections, create `package.xml` files, and save each language's deployment package as a zip file.
- **Preparation**: To download the necessary `.objectTranslation` files from an environment, navigate to **Salesforce Inspector** -> **Download Metadata** -> **ObjectTranslations**. Wait for the download process to complete and then download the files to use them with this tool.

### 6. **Files Comparison**
- **Purpose**: Compares two XLIFF files (old and new) to identify differences in translations.
- **Usage**: Select an old XLIFF file and a new XLIFF file. The tool generates a comparison report in Excel format, highlighting translations that are new, modified, or deleted.

## Version

- **1.2**
