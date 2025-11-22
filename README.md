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

5. **Place your PDF file:**
   - Put your AEON bank statement PDF in the `statement_folder/` directory
   - Update the `input_pdf` path in `statement_converter.py` if your PDF has a different name

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
- Extract tables from the PDF
- Process and clean the transaction data
- Save output to:
  - `AEON_Statement_Extract.xlsx` (Excel format)
  - `AEON_Statement_Extract.csv` (CSV format)

## Configuration

Edit the following variables in `statement_converter.py`:
- `input_pdf`: Path to your PDF file
- `output_excel`: Output Excel filename
- `output_csv`: Output CSV filename

