# ORCID Data Extractor and Report Generator

![Docs](https://img.shields.io/badge/docs-passing-brightgreen.svg)
![Python](https://img.shields.io/badge/python-3.9-blue.svg)
[![PyPI version](https://badge.fury.io/py/your-package-name.svg)](https://pypi.org/project/OrcidXtract/)
[![GitHub repo status](https://img.shields.io/badge/status-maintained-brightgreen.svg)](#)
[![License: MIT](https://img.shields.io/github/license/SafialIslam302/orcidXtractor)](https://github.com/SafialIslam302/orcidXtractor/blob/master/LICENSE.txt)
[![Issues](https://img.shields.io/github/issues/SafialIslam302/orcidXtractor)](https://github.com/SafialIslam302/orcidXtractor/issues)

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

1. Clone the repository:
   ```bash
   git clone https://github.com/SafialIslam302/orcidXtractor.git
   cd orcidXtractor
   
2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt

## Usage

To run the project:

1. Extract ORCID data and generate reports:
   ```bash
   python src/OrcidXtract/main.py --orcid_ids 0000-0000-0000-0000 --output-format txt pdf json --report csv
  
2. Run tests:
   ```bash
   python src/OrcidXtract/test.py

For more detailed information about using the OrcidXtract library, please refer to:

[OrcidXtract on PyPI](https://pypi.org/project/OrcidXtract/)

[Detailed Library README](https://github.com/SafialIslam302/orcidXtractor/blob/master/src/README.md)

## Features

- **ORCID Data Extraction**: Extracts detailed information from ORCID IDs.
- **Multiple Output Formats**: Supports TXT, PDF, JSON, CSV, and Excel formats.
- **Customizable Reports**: Generate reports based on specific requirements.
- **Error Handling**: Handles missing or invalid data gracefully.
- **Automatic Output Handling**: All reports are saved in the `Result` folder.

## License

This project is licensed under the MIT License. See the [LICENSE](https://github.com/SafialIslam302/orcidXtractor/blob/master/LICENSE.txt) file for details.

## Contributions

For any questions or issues, please open an issue on the GitHub repository. If you'd like to contribute, please fork the repository and submit a pull request.
