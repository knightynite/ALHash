# ALHash
Excel Hash Extractor
Overview
The Excel Hash Extractor is a command-line tool that automates the extraction of sheet protection data from Excel files (.xlsx). It scans a specified directory for Excel files, allows the user to select one or more files for processing, and outputs the extracted sheet protection data in a standardized format suitable for security analysis and other auditing purposes.

Features
Automated Discovery: Automatically lists all Excel files in a specified directory.
User Selection Interface: Offers a simple command-line interface for selecting files to process.
Data Extraction and Conversion: Extracts sheet protection data from selected Excel files and converts it into a standardized format.
Batch Processing: Allows for the selection and processing of multiple Excel files in one go.
Installation
This script requires Python and has no external dependencies beyond the standard library. Ensure you have Python 3.x installed on your system to run the script.

Usage
Navigate to the directory containing the ExcelHashExtractor.py script.
Run the script in your terminal or command prompt:
bash
Copy code
python ExcelHashExtractor.py
Follow the on-screen prompts to select Excel files for processing. The script will list all Excel files in the current directory (or a specified directory).
The script will process the selected files and output the extracted sheet protection data in the same directory as the Excel file, within a folder named <ExcelFileName>_SheetProtectionData.
How It Works
The script leverages the zipfile and xml.etree.ElementTree modules to parse the Excel file format (.xlsx), which is essentially a ZIP archive of XML files. It specifically targets the sheet protection data embedded within the workbook's XML structure, extracts this data, and then formats it into a standardized format recognizable for security analysis purposes.

Limitations
The tool is designed for .xlsx files only and does not support older .xls file formats.
It extracts only the sheet protection data, ignoring other forms of protection or encryption that might be present in an Excel file.
Contributing
Contributions are welcome. If you have suggestions for improvement, bug fixes, or new features, feel free to fork the repository and submit a pull request.

License
This project is open source and freely available under the MIT License.


