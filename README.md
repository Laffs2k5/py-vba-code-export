<!-- cSpell:ignore xlsm docm pptm oletools pyproject pytest OOXML venv -->
# VBA Code Exporter

A Python console application to export VBA code from macro-enabled Microsoft Office files (OOXML format: .xlsm, .docm, .pptm).

Extract VBA macros from Office files and save them as separate text files (.bas, .cls, .frm) that can be imported back into Office applications or tracked in version control systems.

## Features

- 🔍 **Extract VBA macros** from Excel (.xlsm), Word (.docm), and PowerPoint (.pptm) files
- 📁 **Automatic file naming** based on module types (class modules, standard modules, forms)
- 🎯 **Intelligent extension detection** for ThisWorkbook, Sheet modules, and custom modules
- 💾 **Export to custom directories** or use the default `exported_vba/` folder
- 🔧 **Command-line interface** for easy automation and scripting
- ✅ **Well-tested** with comprehensive test coverage

## Requirements

- Python 3.8 or higher
- Dependencies managed via `requirements.txt`

## Installation

### Basic Installation

1. Clone the repository:

```bash
git clone <repository-url>
cd py-vba-code-export
```

1. Create a virtual environment (recommended):

```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

1. Install the package:

```bash
pip install -e .
```

### Development Installation

For development with testing tools:

```bash
pip install -e ".[dev]"
```

## Usage

### Basic Usage

Export VBA from an Office file to the default directory (`exported_vba/`):

```bash
vba-export myfile.xlsm
```

### Custom Output Directory

Specify a custom directory for exported files:

```bash
vba-export myfile.xlsm output/vba_modules
```

### Verbose Output

Show detailed progress information:

```bash
vba-export myfile.xlsm -v
```

### Examples

Export VBA from an Excel workbook:

```bash
vba-export budget.xlsm
```

Export from a Word document to a specific folder:

```bash
vba-export report.docm exported_macros/
```

Export with verbose output to see all extracted modules:

```bash
vba-export presentation.pptm --verbose
```

### Supported File Formats

- **Excel**: `.xlsm` (macro-enabled workbooks)
- **Word**: `.docm` (macro-enabled documents)
- **PowerPoint**: `.pptm` (macro-enabled presentations)

> **Note**: This tool works with OOXML format files only, not legacy binary formats (.xls, .doc, .ppt)

## Output

The tool creates separate files for each VBA module with appropriate extensions:

- `.bas` - Standard modules
- `.cls` - Class modules (including ThisWorkbook and Sheet modules)
- `.frm` - UserForm modules

**Example output:**

```text
exported_vba/
├── ThisWorkbook.cls
├── Sheet1.cls
├── Module1.bas
├── UserForm1.frm
└── MyClass.cls
```

## Development

### Running Tests

Run all tests:

```bash
pytest
```

Run with verbose output:

```bash
pytest -v
```

Run with coverage report:

```bash
pytest --cov=src/vba_export --cov-report=term-missing
```

### Project Structure

```text
py-vba-code-export/
├── src/
│   └── vba_export/
│       ├── __init__.py
│       ├── cli.py          # Command-line interface
│       └── exporter.py     # Core extraction logic
├── tests/
│   ├── test_cli.py         # CLI tests
│   └── test_exporter.py    # Exporter tests
├── pyproject.toml          # Project configuration
├── requirements.txt        # Dependencies
└── README.md
```

## How It Works

The tool uses the [oletools](https://github.com/decalage2/oletools) library to parse Office files and extract VBA macros. It:

1. Opens the macro-enabled Office file
2. Extracts all VBA modules using `VBA_Parser`
3. Determines the appropriate file extension based on module type
4. Saves each module as a separate text file with proper encoding

## Troubleshooting

### Command not found

If `vba-export` is not found after installation, make sure:

- You've activated your virtual environment
- The package was installed with `pip install -e .`
- Your Python scripts directory is in your PATH

### No modules found

If the tool reports no VBA modules found:

- Verify the file is actually macro-enabled (.xlsm, .docm, or .pptm)
- Ensure the file contains VBA macros
- Try opening the file in Office and checking the VBA editor (Alt+F11)

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit issues or pull requests.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes and add tests
4. Ensure all tests pass (`pytest`)
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request
