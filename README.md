# Financial PDF Parser

A Python-based tool for extracting structured data from financial PDF documents using predefined templates. The extracted data is formatted and saved to Excel files with color-coded table names and headers.

## Features

- Extract tables from PDF files using Tabula templates
- Support for multiple templates processing
- Excel output with formatted styling:
  - Table names highlighted in green
  - Column headers highlighted in yellow
  - 5-row gaps between tables for better readability
- Cross-platform support (Windows & Ubuntu/Linux)
- Automatic data cleaning and formatting

## Prerequisites

### System Requirements

#### Ubuntu/Linux
```bash
# Update package list
sudo apt update

# Install Java (required for Tabula)
sudo apt install default-jdk

# Verify Java installation
java -version

# Install Python 3 and pip (if not already installed)
sudo apt install python3 python3-pip

# Optional: Install Python virtual environment
sudo apt install python3-venv
```

#### Windows
1. **Install Java:**
   - Download Java from [Oracle](https://www.oracle.com/java/technologies/downloads/) or [OpenJDK](https://adoptium.net/)
   - Install and add Java to your PATH
   - Verify installation by opening Command Prompt and running: `java -version`

2. **Install Python:**
   - Download Python 3.8+ from [python.org](https://www.python.org/downloads/)
   - During installation, check "Add Python to PATH"
   - Verify installation: `python --version`

### Python Dependencies

The project requires the following Python packages:
- `tabula-py` - PDF table extraction
- `pandas` - Data manipulation
- `openpyxl` - Excel file handling
- `pathlib` - Path handling (built-in)

## Installation

### 1. Clone the Repository
```bash
git clone https://github.com/AliTahir-101/fi-financial-pdf-parser.git
cd fi-financial-pdf-parser
```

### 2. Set Up Virtual Environment (Recommended)

#### Ubuntu/Linux
```bash
# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate
```

#### Windows
```cmd
# Create virtual environment
python -m venv venv

# Activate virtual environment
venv\Scripts\activate
```

### 3. Install Python Dependencies

```bash
# Install required packages
pip install tabula-py pandas openpyxl

# Or install from requirements file (if you create one)
pip install -r requirements.txt
```

## Project Structure

```
fi-financial-pdf-parser/
├── fi_financial_pdf_parser.py    # Main script
├── input_pdf_files/              # Input PDF files directory
│   ├── new_format/
│   │   └── test.pdf
│   └── old_format/
│       └── test.pdf
├── templates/                    # Tabula template files
│   ├── new_format/
│   │   ├── 1_table_name.tabula-template.json
│   │   ├── 1_table_data.tabula-template.json
│   │   ├── 2_table_name.tabula-template.json
│   │   ├── 2_table_data.tabula-template.json
│   │   └── ...
│   └── old_format/
│       └── old_format.tabula-template.json
├── output.xlsx                   # Generated output file
└── README.md                     # This file
```

## Usage

### Basic Usage

1. **Place your PDF file** in the `input_pdf_files/new_format/` directory
2. **Prepare your templates** in the `templates/new_format/` directory
3. **Configure the script** by editing the main configuration in `fi_financial_pdf_parser.py`:

```python
def main():
    # Configuration - change these paths as needed
    pdf_file = "./input_pdf_files/new_format/test.pdf"
    template_dir = "./templates/new_format"
    output_file = "./output.xlsx"
    num_templates = 5  # Define how many templates to process
```

4. **Run the script:**

#### Ubuntu/Linux
```bash
python3 fi_financial_pdf_parser.py
```

#### Windows
```cmd
python fi_financial_pdf_parser.py
```

### Template Structure Options

The script supports two template naming conventions:

#### Option 1: Separate Name and Data Templates
- `1_table_name.tabula-template.json` - For extracting table names
- `1_table_data.tabula-template.json` - For extracting table data
- `2_table_name.tabula-template.json`
- `2_table_data.tabula-template.json`
- And so on...

#### Option 2: Single Numbered Templates
- `1.tabula-template.json`
- `2.tabula-template.json`
- `3.tabula-template.json`
- And so on...

### Creating Tabula Templates

1. **Install Tabula GUI** (optional but recommended for template creation):
   - Download from [tabula.technology](https://tabula.technology/)
   
2. **Create templates using Tabula GUI:**
   - Open your PDF in Tabula
   - Select the table area you want to extract
   - Save the template as JSON
   - Place the template file in the appropriate directory

3. **Template naming:**
   - Follow the naming convention described above
   - Ensure template numbers match your `num_templates` configuration

## Configuration Options

### Main Configuration Parameters

```python
# In the main() function of fi_financial_pdf_parser.py

pdf_file = "./input_pdf_files/new_format/test.pdf"  # Path to your PDF
template_dir = "./templates/new_format"              # Template directory
output_file = "./output.xlsx"                       # Output Excel file
num_templates = 5                                   # Number of templates to process
```

### Template Processing

- The script processes templates numbered from 1 to `num_templates`
- If a template file is missing, it will be skipped with a warning
- The script automatically falls back to single template files if separate name/data templates are not found

## Output

The script generates an Excel file with the following features:

- **Multiple tables** from different templates in a single worksheet
- **Green highlighting** for table names
- **Yellow highlighting** for column headers
- **5-row gaps** between tables for visual separation
- **Auto-sized columns** for better readability
- **Clean data** with removed empty rows/columns and trimmed whitespace

## Troubleshooting

### Common Issues

1. **Java not found error:**
   ```
   Error: Java not found. Please install Java and add it to your PATH.
   ```
   **Solution:** Install Java and ensure it's in your system PATH.

2. **Template file not found:**
   ```
   Table data template not found: ./templates/new_format/1_table_data.tabula-template.json
   ```
   **Solution:** Check template file names and paths. Ensure they follow the expected naming convention.

3. **PDF file not found:**
   ```
   PDF file not found: ./input_pdf_files/new_format/test.pdf
   ```
   **Solution:** Place your PDF file in the correct directory or update the `pdf_file` path in the configuration.

4. **Permission errors on Windows:**
   **Solution:** Run Command Prompt as Administrator or check file permissions.

### Debug Tips

- Check the console output for detailed processing information
- Verify that all required directories exist
- Ensure PDF files are not password-protected
- Test with a single template first before processing multiple templates

## Requirements File

Create a `requirements.txt` file for easy dependency management:

```txt
tabula-py>=2.5.1
pandas>=1.5.0
openpyxl>=3.0.10
```

Install dependencies using:
```bash
pip install -r requirements.txt
```

## Testing

### Quick Test

1. Ensure you have a test PDF in `input_pdf_files/new_format/test.pdf`
2. Create at least one template file in `templates/new_format/`
3. Set `num_templates = 1` in the main configuration
4. Run the script and check the generated `output.xlsx`

### Validation

- Open the generated Excel file
- Verify that table names are highlighted in green
- Verify that column headers are highlighted in yellow
- Check that there are 5-row gaps between tables (if multiple tables)
- Ensure data is properly formatted and cleaned

## License

[Add your license information here]

## Contributing

[Add contribution guidelines here]

## Support

For issues and questions, please [create an issue](https://github.com/AliTahir-101/fi-financial-pdf-parser/issues) on GitHub.