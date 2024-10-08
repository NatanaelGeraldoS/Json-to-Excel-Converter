# BEBEK CONVERTER

BEBEK CONVERTER is a simple Python application that converts JSON files into either Excel (.xlsx) or CSV format. The application also provides options to handle questionnaire data specifically, with features like bold headers and auto-fitting columns in Excel.

## Features

- **Convert JSON to Excel or CSV**: Easily convert your JSON data to a more readable format.
- **Questionnaire Support**: Special handling for questionnaire data, including metadata and structured output.
- **Excel Customization**: Choose to bold headers and auto-fit columns when saving to Excel format.
- **User-Friendly GUI**: A straightforward graphical interface built with `tkinter`.

## Requirements

- Python 3.7 or higher
- Required Python libraries:
  - `tkinter`
  - `pandas`
  - `openpyxl`

You can install the required Python libraries using the following command:

```bash
pip install pandas openpyxl
```

## Installation
1. Clone the Repository:

```bash
git clone https://github.com/NatanaelGeraldoS/Json-to-Excel-Converter
```

2. Navigate to the Project Directory:

```bash
cd bebek-converter
```
3. Install the Dependencies:
Install the required Python libraries if you haven't already:

```bash
pip install -r requirements.txt
```
4. Run the Application:

```bash
python bebek_converter.py
```
## Usage
1. Upload a JSON File:
    - Click on the "Browse" button to select the JSON file you want to convert.
2. Select Output Format:
    - Choose between Excel or CSV format.
3. Customize Output (Optional):
    - Select "Bold Headers" to bold the headers in the Excel file.
    - Select "Auto Fit Columns" to automatically adjust the column widths in the Excel file.
    - Select "Is Questionnaire" if you are converting a questionnaire JSON file.
4. Save the Output:
    - Click "Save As" to specify the output file name and location.
5. Convert:
    - Click the "Convert" button to perform the conversion.
6. Success Message:
    - A success message will be displayed when the file is saved successfully.
## Packaging the Application
If you want to distribute the application as a standalone executable:

1. Install PyInstaller:
```bash
pip install pyinstaller]
```
2. Create the Executable:

```bash
pyinstaller --onefile --noconsole --icon=duck.ico  --add-data "duck.png;." --add-data "duck.ico;." main.py
```
3. Distribute:
    - The executable will be located in the dist folder. You can share this executable with others who do not have Python installed.

## Contributing
Contributions are welcome! If you have ideas for improvements or find any issues, please submit a pull request or create an issue.


## Acknowledgements
- tkinter - Python's standard GUI toolkit.
- pandas - Data manipulation and analysis library.
- openpyxl - A library to read/write Excel 2010 xlsx/xlsm/xltx/xltm files.