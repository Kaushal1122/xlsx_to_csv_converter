# XLSX to CSV Converter

A Python command-line application that converts every worksheet in an Excel workbook or Google Sheets spreadsheet into separate CSV files.

## Features

- Supports local `.xlsx` files.
- Supports publicly accessible Google Sheets URLs.
- Converts every worksheet into an individual CSV file.
- Creates an output directory named after the input workbook.
- Handles invalid files and unsupported URLs.
- Configurable download timeout, CSV encoding, and logging.
- Uses logging for application monitoring and debugging.
- Includes unit tests using Pytest.
- Follows Python code quality standards using Ruff.

## Project Structure

```text
xlsx_to_csv_converter/
│
├── converter.py
├── config.ini
├── requirements.txt
├── README.md
├── .gitignore
│
├── docs/
│   └── sample_data_python_assignment.xlsx
│
└── tests/
    ├── __init__.py
    └── test_converter.py
```

## Requirements

- Python 3.10 or later
- pip

## Installation

Clone the repository:

```bash
git clone https://github.com/Kaushal1122/xlsx_to_csv_converter.git
cd xlsx_to_csv_converter
```

Create a virtual environment:

```bash
python -m venv .venv
```

Activate it on Windows:

```powershell
.venv\Scripts\Activate.ps1
```

Install dependencies:

```bash
python -m pip install -r requirements.txt
```

## Usage

### 1. Convert a local Excel file

```bash
python converter.py path/to/workbook.xlsx
```

### 2. Convert a Google Sheets spreadsheet

```bash
python converter.py "https://docs.google.com/spreadsheets/d/1jFpj6RY4hW0tZ56bYepVOYC5AkAIRJYpozgmi4_pzak/edit?usp=sharing"
```

The Google spreadsheet must be publicly accessible or shared with anyone who has the link as a viewer.

**Sample Google Sheets spreadsheet:** [Open the sample spreadsheet](https://docs.google.com/spreadsheets/d/1jFpj6RY4hW0tZ56bYepVOYC5AkAIRJYpozgmi4_pzak/edit?usp=sharing)

### 3. Specify a configuration file

```bash
python converter.py workbook.xlsx --config config.ini
```

## Output

The application creates a directory named after the input workbook.

For example, converting `sales.xlsx` creates:

```text
sales/
├── Office Supply Sales.csv
└── Food Sales.csv
```

Each worksheet is saved as a separate CSV file.

## Configuration

Application settings are maintained in `config.ini`.

```ini
[DOWNLOAD]
timeout = 30

[CSV]
encoding = utf-8

[LOGGING]
level = INFO
file = converter.log
```

| Setting | Description |
|---|---|
| timeout | Google Sheets download timeout in seconds |
| encoding | Character encoding for CSV files |
| level | Logging level |
| file | Log file name |

## Running Tests

Run all unit tests:

```bash
python -m pytest -v
```

## Code Quality

Check code quality using Ruff:

```bash
python -m ruff check .
```

## Technologies Used

- Python
- OpenPyXL
- CSV module
- Requests
- ConfigParser
- Logging
- Pytest
- Ruff
- GitHub Actions

## Assumptions

- Input Excel files must use the `.xlsx` format.
- Each worksheet contains exactly one header row.
- Google Sheets must be accessible without authentication.
- Each worksheet is exported as an individual CSV file.

## Repository

GitHub: [Kaushal1122/xlsx_to_csv_converter](https://github.com/Kaushal1122/xlsx_to_csv_converter)