# Report-Comparison
Python script to compare Excel files from different releases. Reads data from Release1 and Release2 folders and creates a comparison report with separate sheets for each release.

## Project Overview

This tool automates the process of comparing Excel reports from different releases (Release1 vs Release2) and generates a comprehensive comparison report with:
- Separate sheets for Release1 data (Release1_Report)
- Separate sheets for Release2 data (Release2_Report)
- Organized folder structure for easy management
- Error handling and validation

## Folder Structure

```
ProjectFolder/
├── Release1/
│   └── A42/
│       └── excelFile1.xlsx
├── Release2/
│   └── A42/
│       └── excelFile2.xlsx
├── comparison/
│   └── A42/
│       └── comparison_report.xlsx
├── compare_excel_reports.py
├── requirements.txt
└── README.md
```

## Installation

### Step 1: Clone Repository
```bash
git clone https://github.com/somnathsarak/Report-Comparison.git
cd Report-Comparison
```

### Step 2: Create Virtual Environment
```bash
python -m venv venv
venv\Scripts\activate  # Windows
source venv/bin/activate  # macOS/Linux
```

### Step 3: Install Dependencies
```bash
pip install -r requirements.txt
```

## Required Libraries

- pandas >= 1.3.0
- openpyxl >= 3.6.0
- xlrd >= 2.0.1

## Usage

1. Place Excel files in:
   - Release1/A42/excelFile1.xlsx
   - Release2/A42/excelFile2.xlsx

2. Run script:
```bash
python compare_excel_reports.py
```

3. Output generated at: comparison/A42/comparison_report.xlsx

## Features

- Reads Excel data from Release1 and Release2
- Creates comparison report with separate sheets
- Automatic folder structure creation
- Comprehensive error handling
- Non-destructive (original files untouched)

## Troubleshooting

If you encounter module import errors:
```bash
pip install -r requirements.txt --upgrade
```

For file not found errors, ensure Excel files are in correct paths.

