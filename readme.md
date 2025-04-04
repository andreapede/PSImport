# PSImport - Palmsense PStouch Data Converter ![Python](https://img.shields.io/badge/python-3.8%2B-blue) ![License](https://img.shields.io/badge/license-MIT-green)

📦 **PSImport** is a Python tool for importing, processing, and converting cyclic voltammetry data from CSV files exported by the Palmsense PStouch software.

## 🚀 Overview

PSImport is specifically designed to handle the unique format of CSV files generated by Palmsense PStouch electrochemical analysis software. It extracts voltammetry data from these files and converts it into various formats suitable for analysis.

## 🌟 Features

- **Automatic format detection**: Handles the UTF-16 encoding and specific structure of Palmsense exports
- **Multiple scan support**: Correctly separates and processes multiple scans within a single file
- **Metadata extraction**: Preserves measurement dates, scan names, and other metadata
- **Multiple export formats**:
  - Excel (.xlsx) with separate sheets for each scan
  - CSV for simple data sharing
  - TXT for universal compatibility
  - CHI format for CH Instruments compatibility

## 🛠️ Installation

### Requirements

Install the required Python packages:

```bash
pip install numpy pandas openpyxl
```

### Getting Started

Clone this repository or download the `psimport.py` file:

```bash
git clone https://github.com/username/psimport.git
cd psimport
```

### 📚 Usage

#### Command Line Interface

Run the script with the following command:

```bash
python psimport.py input_file.csv [options]
```

#### Options

- `-o`, `--output`: Output file path (default: `output.xlsx`)
- `-f`, `--format`: Output format - choose from `csv`, `excel`, `txt`, `chi` (default: `excel`)
- `-s`, `--scan`: Index of scan to export (0-based, default: `0`)
- `-e`, `--encoding`: Input file encoding (default: `utf-16`)

#### Examples

Export all scans to Excel:

```bash
python psimport.py "sample.csv" -o voltammetry.xlsx
```

Export a specific scan to CSV:

```bash
python psimport.py "sample.csv" -o scan2.csv -f csv -s 1
```

Export to TXT format:

```bash
python psimport.py "sample.csv" -o data.txt -f txt -s 0
```

Export to CHI format:

```bash
python psimport.py "sample.csv" -o data.chi -f chi -s 0
```

#### Python API

You can use PSImport as a module in your own Python scripts:

```python
from psimport import VoltammetryImporter

# Initialize and load data
importer = VoltammetryImporter("sample.csv")
if importer.load_file():
    # Get information about scans
    print(f"Number of scans: {importer.get_scan_count()}")
    print(f"Scan names: {importer.get_scan_names()}")

    # Get data for a specific scan
    scan_data = importer.get_scan_data(0)
    print(f"First scan: {scan_data['name']}")
    print(f"Points: {len(scan_data['potential'])}")

    # Export in various formats
    importer.export_to_excel("results.xlsx")
    importer.export_to_csv("data.csv", scan_index=None)  # All scans
    importer.export_to_txt("scan1.txt", scan_index=0)
    importer.export_to_chi("data.chi", scan_index=0)
```

### 📋 Output Format Details

#### Excel (.xlsx)

- Separate worksheet for each scan
- Summary worksheet with all scans side-by-side
- Metadata sheet with file information
- Headers and metadata within each scan sheet

#### CSV (.csv)

- Single scan export: Two columns (Potential, Current)
- All scans export: Multiple columns with scan identifiers

#### TXT (.txt)

- Tab-separated columns
- Simple header
- Compatible with most data analysis software

#### CHI (.chi)

- Format compatible with CH Instruments software
- Proper headers with technique information
- Current values converted from µA to A as required

## ⚙️ Technical Details

PSImport is designed to handle the specific format of Palmsense PStouch exports:

- UTF-16 encoded CSV files
- Multiple scan data in parallel columns
- Scan headers in specific rows
- Potential (V) and current (µA) paired columns

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

Palmsense for their PStouch electrochemical analysis software
Contributors and users who provide feedback and suggestions
