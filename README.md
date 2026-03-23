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



## A42 Allowance Report Combiner

### Overview

The **A42_combine_allowance_reports.py** script is specifically designed to combine Allowance Delivery Reports from two different releases (R3.6 and R3.7) into a single Excel file with properly named sheets.

### Updated Folder Structure

The project now supports the following organized folder structure:

```
ProjectFolder/
├── Release1/
│   └── A42/
│       └── Allowance-Delivery-Report-POST_R3.6_A42_CA_QC_JOINT_AUCTION-07-01-02-01-07-2026.xlsx
├── Release2/
│   └── A42/
│       └── Allowance-Delivery-Report-R3.7_A42_CA_QC_JOINT_AUCTION-6March_01-03-06-2026-2.xlsx
├── comparison/
│   └── A42/
│       ├── A42_combine_allowance_reports.py
│       └── allowance_comparison_report.xlsx (Generated output)
├── combine_allowance_reports.py
├── compare_excel_reports.py
├── requirements.txt
└── README.md
```

### How to Use A42_combine_allowance_reports.py

#### Step 1: Navigate to the Comparison Directory

```bash
cd comparison/A42/
```

#### Step 2: Run the Script

```bash
python A42_combine_allowance_reports.py
```

The script will automatically:
1. Navigate to the project root
2. Read R3.6 file from: `Release1/A42/`
3. Read R3.7 file from: `Release2/A42/`
4. Create comparison report in: `comparison/A42/allowance_comparison_report.xlsx`

#### Step 3: Specify Custom Filenames (Optional)

If your filenames differ from the defaults, pass them as arguments:

```bash
python A42_combine_allowance_reports.py "your_r36_filename.xlsx" "your_r37_filename.xlsx"
```

### Output Details

The generated Excel file contains:

- **Sheet 1**: `POST_R3.6_A42_CA_QC_JOINT_AUCTION` - Contains all R3.6 data
- **Sheet 2**: `R3.7_A42_CA_QC_JOINT_AUCTION` - Contains all R3.7 data

Both sheets include:
- Headers from the source files
- All data rows
- Proper formatting and data types

### Key Features

✅ **Automatic Path Resolution**: Works from `comparison/A42/` directory
✅ **Error Handling**: Validates file existence and provides clear error messages
✅ **Customizable**: Supports command-line arguments for flexible usage
✅ **Detailed Logging**: Shows step-by-step progress and final summary
✅ **Non-Destructive**: Original files remain untouched
✅ **Report Summary**: Displays row and column counts for verification

### Example Output

```
======================================================================
ALLOWANCE DELIVERY REPORT COMBINER
======================================================================

Project Directory: /path/to/ProjectFolder
Output Directory: current

Reading Release 1 (R3.6) file...
  Path: /path/to/Release1/A42/Allowance-Delivery-Report-POST_R3.6_A42_CA_QC_JOINT_AUCTION-07-01-02-01-07-2026.xlsx
Successfully read 1 sheet(s) from: Allowance-Delivery-Report-POST_R3.6_A42_CA_QC_JOINT_AUCTION-07-01-02-01-07-2026.xlsx

Reading Release 2 (R3.7) file...
  Path: /path/to/Release2/A42/Allowance-Delivery-Report-R3.7_A42_CA_QC_JOINT_AUCTION-6March_01-03-06-2026-2.xlsx
Successfully read 1 sheet(s) from: Allowance-Delivery-Report-R3.7_A42_CA_QC_JOINT_AUCTION-6March_01-03-06-2026-2.xlsx

Creating sheets in output workbook...
  Sheet 1: POST_R3.6_A42_CA_QC_JOINT_AUCTION
  Sheet 2: R3.7_A42_CA_QC_JOINT_AUCTION

======================================================================
SUCCESS: Combined report saved to: allowance_comparison_report.xlsx
======================================================================

Report Summary:
  Sheet 1 (POST_R3.6_A42_CA_QC_JOINT_AUCTION): 15 rows, 6 columns
  Sheet 2 (R3.7_A42_CA_QC_JOINT_AUCTION): 16 rows, 6 columns
  Total sheets in output: 2

Output file location: /path/to/comparison/A42/allowance_comparison_report.xlsx

✓ Report comparison completed successfully!
```

### Troubleshooting

**Issue**: "FileNotFoundError: File not found"

**Solution**: Ensure your Excel files are in the correct locations:
```
Release1/A42/[your_r36_filename].xlsx
Release2/A42/[your_r37_filename].xlsx
```

**Issue**: Module import errors

**Solution**: Make sure dependencies are installed:
```bash
pip install -r ../../requirements.txt
```

### Directory Navigation

If you run the script from the project root instead:
```bash
python comparison/A42/A42_combine_allowance_reports.py
```

The script handles path resolution automatically.
