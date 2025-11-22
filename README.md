# Bank Statement Converter

Converts AEON bank statement PDFs to Excel and CSV formats.

## Setup

1. **Create and activate a virtual environment (recommended):**
   
   **Windows:**
   ```powershell
   python -m venv venv
   .\venv\Scripts\Activate.ps1
   ```
   
   **Linux/Mac:**
   ```bash
   python -m venv venv
   source venv/bin/activate
   ```

2. **Upgrade pip (recommended):**
   ```powershell
   python -m pip install --upgrade pip
   ```

3. **Install Python dependencies:**
   ```powershell
   pip install -r requirements.txt
   ```
   
   **If you encounter compilation errors**, try one of these solutions:
   
   **Option 1: Force using pre-built wheels (recommended):**
   ```powershell
   pip install -r requirements.txt --only-binary :all: --no-cache-dir
   ```
   
   **Option 2: Install in steps:**
   ```powershell
   pip install numpy pandas openpyxl --only-binary :all:
   pip install opencv-python camelot-py --only-binary :all:
   ```

4. **Install system dependencies (required for camelot):**
   - **Windows:** Download and install [Ghostscript](https://www.ghostscript.com/download/gsdnld.html)
   - Make sure Ghostscript is added to your system PATH
   - **Note:** If Ghostscript is not available, the script will automatically fall back to using pdfplumber

5. **Place your PDF files:**
   - Put your AEON bank statement PDFs in the `statement_folder/` directory
   - The script will automatically process all PDF files in this folder

## Usage

**Make sure your virtual environment is activated**, then run the script:
```bash
python statement_converter.py
```

**To deactivate the virtual environment when done:**
```bash
deactivate
```

The script will:
- Automatically find and process all PDF files in the `statement_folder/` directory
- Extract tables from each PDF using camelot (with pdfplumber as fallback)
- Process and clean the transaction data (handles multi-line descriptions, merges transaction date rows, removes duplicates)
- Save output files to the `processed_output/` folder
- Generate output filenames automatically based on the PDF filename (e.g., `1000073282_202501.xlsx` and `1000073282_202501.csv`)

## Configuration

You can modify these variables in `statement_converter.py` if needed:
- `statement_folder`: Folder containing input PDF files (default: `"statement_folder"`)
- `output_folder`: Folder for output files (default: `"processed_output"`)

