# ORCID Data Extractor and Report Generator

This Python script extracts ORCID information from a file containing ORCID IDs and generates reports in various formats (TXT, PDF, JSON, CSV, Excel).

## Table of Contents
- [Installation](#installation)
- [Usage](#usage)
  - [Command-Line Usage](#command-line-usage)
  - [Using in Python Code](#using-in-python-code)
- [Features](#features)
- [Dependencies](#dependencies)
- [License](#license)
- [Contributions](#contributions)

## Installation

1. To install the `OrcidXtract` library, simply use `pip`:
   ```bash
   pip install OrcidXtract

## Usage

Once the library is installed, you can use it from the command line with the `orcidxtract` command.

### Command-Line Arguments

      orcidxtract --help
   
      Options:
        -h, --help            show this help message and exit
        --inputfile INPUTFILE
                              Path to the input file containing ORCID IDs.
        --orcid_ids [ORCID_IDS ...]
                              List of ORCID IDs (e.g., 0000-0001-...).
        --output-format {txt,pdf,json} [{txt,pdf,json} ...]
                              Specify one or more output formats (txt, pdf, json).
                              Example: --output-format txt pdf json
        --report {csv,excel}  Specify if you want to generate a CSV or Excel report.
                              Example: --report csv

The script accepts the following command-line arguments:

- `--inputfile`: Path to the input file containing ORCID IDs.
- `--orcid_ids`: A list of ORCID IDs (e.g., `0000-0001-5109-3700`).
- `--output-format`: Specify one or more output formats (txt, pdf, json). Example: `--output-format txt pdf json`. 
- `--report`: Specify if you want to generate a CSV or Excel report.

#### Example Commands

1. Extract ORCID data and generate TXT, PDF, and JSON reports:
   ```bash
   orcidxtract --inputfile orcid_ids.txt --output-format txt pdf json
2. Extract ORCID data and generate a CSV report:
   ```bash
   orcidxtract --inputfile orcid_ids.txt --output-format txt --report csv
3. Extract ORCID data and generate an Excel report:
   ```bash
   orcidxtract --inputfile orcid_ids.txt --output-format json --report excel
4. Extract ORCID directly from the command line (Work only for one ORCID id at a time):
   ```bash
   orcidxtract --orcid_ids xxxx-xxxx-xxxx-xxxx --output-format txt pdf json  --report csv   
5. Generate only Report
   ```bash
   orcidxtract --orcid_ids xxxx-xxxx-xxxx-xxxx --report csv

#### Function Details

**Fetching ORCID Data**: `get_orcid_data(orcid_id: str)`

This function fetches ORCID profile information using an ORCID ID.

   **Parameters**:
   - `orcid_id` (*str*): The ORCID ID of the researcher.
   
   **Returns**:
   - An `OrcidData` object containing ORCID profile details.

**Generate File**: `create_report(orcid_data: list, format: str)`

he library provides multiple functions to generate files in different formats: `TXT`, `PDF`, and `JSON`.

   **Functions:**
   - `create_txt(file_path: str, orcid_info: OrcidData)`
   - `create_pdf(file_path: str, orcid_info: OrcidData)`
   - `create_json(file_path: str, orcid_info: OrcidData)`

   **Parameters**:
   - `file_path` (*str*): The output path where the JSON file should be saved. 
   - `orcid_info` (*OrcidData*): The ORCID data object.

**Generating Reports**: `create_report(orcid_data: list, format: str)`

Generates a report in CSV or Excel format, containing multiple ORCID profiles.

   **Parameters**:
   - `orcid_data` (*list*): A list of OrcidData objects.
   - `format` (*str*): The format of the report (csv or excel).

#### Using in Python Code

To integrate the library into a Python script, follow these steps.

1. **Import the Required Functions**

   ```bash
   from OrcidXtract.orcid_extractor import get_orcid_data
   from OrcidXtract.report_generator import create_txt, create_pdf, create_json, create_report

2. **Define the ORCID IDs**

   You can manually define ORCID IDs or read them from a file.

   ```bash
   # Define ORCID IDs to fetch data for
   orcid_ids = ["0000-0002-2140-1538", "0009-0003-9071-3804"]

   # Read ORCID IDs from a file
   input_file = "input_file.txt"
   with open(input_file, "r") as f:
         orcid_ids = [line.strip() for line in f if line.strip()]
   
3. **Set Output Formats**

   Specify the output formats (txt, pdf, json) and report type (csv or excel).
   ```bash
   output_formats = ["txt", "json", "pdf"]  # Choose desired formats
   report_format = "csv"  # Change to "excel" if needed

4. **Fetch ORCID Data and Generate Reports**

   ```bash
   orcid_data = []
   for orcid_id in orcid_ids:
       orcid_info = get_orcid_data(orcid_id)
       orcid_data.append(orcid_info)

       # Save reports only in the specified formats
       if "txt" in output_formats:
           create_txt(f"{orcid_id}.txt", orcid_info)
       if "pdf" in output_formats:
           create_pdf(f"{orcid_id}.pdf", orcid_info)
       if "json" in output_formats:
           create_json(f"{orcid_id}.json", orcid_info)

5. **Generate a Summary Report (CSV/Excel)**

   ```bash
   create_report(orcid_data, report_format)  # Generates a CSV or Excel report

### Input File Format

The input file should contain one ORCID ID per line. Example (`orcid_ids.txt`):

    0000-0000-0000-0001
    0000-0000-0000-0002
    0000-0000-0000-0003

### Output Files

The generated reports will be saved in the **Result** directory. The directory structure will look like this:

    Result
      - 0000-0000-0000-0001.txt
      - 0000-0000-0000-0002.pdf
      - 0000-0000-0000-0003.json
      - orcid_report.csv
      - orcid_report.xlsx

## Features

- **ORCID Data Extraction**: Extracts detailed information from ORCID IDs.
- **Multiple Output Formats**: Supports TXT, PDF, JSON, CSV, and Excel formats.
- **Customizable Reports**: Generate reports based on specific requirements.
- **Error Handling**: Handles missing or invalid data gracefully.
- **Automatic Output Handling**: All reports are saved in the `Result` folder.

## Dependencies

- **argparse**: For parsing command-line arguments.
- **reportlab**: For generating PDF reports.
- **pandas**: For generating CSV and Excel reports.
- **requests**: For making HTTP requests to ORCID's API.
- **openpyxl**: For working with Excel files.

These dependencies will be automatically installed when you install the package via pip.

## License

This project is licensed under the MIT License. See the [LICENSE](https://github.com/SafialIslam302/ORCID-Information/blob/master/LICENSE) file for details.

## Contributions

For any questions or issues, please open an issue on the GitHub repository. If you'd like to contribute, please fork the repository and submit a pull request.
