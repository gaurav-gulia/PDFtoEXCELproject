# PDF To Excel sheet Extractor

This tool extracts tables from PDF files and saves them into an Excel file. The tool leverages `PyMuPDF` for PDF processing, `pandas` for data manipulation, and `openpyxl` for Excel file handling.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Dependencies](#dependencies)
- [License](#license)

## Installation

1. Clone the repository or download the source code.
   
2. Install the required dependencies:
    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. Place the PDF file you want to extract tables from in the project directory.

2. Modify the `main` function in `main.py` to specify the input PDF file path and the desired output Excel file path:
    ```python
    pdf_path = "test9.pdf"
    extracted_excel_path = "extracted_tables.xlsx"
    ```

3. Run the script:
    ```bash
    python main.py
    ```

4. The extracted tables will be saved in the specified Excel file.

## Dependencies

The project requires the following Python libraries:

- `pymupdf`
- `pandas`
- `openpyxl`

These dependencies are listed in the `requirements.txt` file.


## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.
