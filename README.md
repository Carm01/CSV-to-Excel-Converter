# CSV to Excel Converter

This Python script is a simple, no-frills converter that transforms CSV or tab-delimited text files into Excel (.xlsx) format. It handles Unicode characters, ensuring proper encoding for diverse datasets. The script includes no graphical user interface (GUI) beyond basic file picker dialogs for selecting input and output files. Below are instructions to set up and run the script, including how to run it using the native Python IDLE editor via right-click, and how to compile it into an executable (.exe) file without a console window.

## Prerequisites

### 1. Install Python
- **Download Python**: Get Python 3.8 or later from the [official Python website](https://www.python.org/downloads/). Choose the appropriate version for your operating system (Windows, macOS, or Linux).
- **Installation**:
  - **Windows**: Run the installer, check "Add Python to PATH," and select "Install Now." Ensure the option to install `py launcher` and IDLE is selected (default in most installers).
  - **macOS/Linux**: Follow the installer instructions or use a package manager like Homebrew (`brew install python` on macOS) or apt (`sudo apt install python3` on Ubuntu). IDLE is included with standard Python installations.
- **Verify Installation**: Open a terminal or command prompt and run:
  ```bash
  python --version
  ```
  or
  ```bash
  python3 --version
  ```
  Ensure the version is 3.8 or higher. To verify IDLE, run:
  ```bash
  python -m idlelib
  ```
  This should open the IDLE editor.

### 2. Install Required Modules
The script requires the `xlsxwriter` module. The `tkinter` and `csv` modules are included with Python by default.

- **Install `xlsxwriter`**:
  Open a terminal or command prompt and run:
  ```bash
  pip install xlsxwriter
  ```
- **Verify Installation**:
  Run the following in Python (e.g., in IDLE) to check if the module is installed:
  ```python
  import xlsxwriter
  print(xlsxwriter.__version__)
  ```

## Running the Script

### Option 1: Run via Terminal
1. **Save the Script**:
   - Copy the script into a file named `csv_to_excel.py` (or any preferred name).
2. **Execute the Script**:
   - Open a terminal or command prompt in the script's directory.
   - Run:
     ```bash
     python csv_to_excel.py
     ```
   - File picker dialogs will prompt you to select an input file (.csv, .txt, .tsv, or .xls) and specify an output .xlsx file.

### Option 2: Run via Right-Click with IDLE (Windows)
1. **Ensure Python and IDLE are Associated with .py Files**:
   - During Python installation, the `py launcher` and IDLE should associate `.py` files with Python. To verify, right-click `csv_to_excel.py`, select "Open with," and ensure "IDLE" is listed. If not, choose "Open with" > "Choose another app" > select `idle.bat` (typically in `C:\PythonXX\Lib\idlelib\idle.bat` or `C:\Users\<YourUser>\AppData\Local\Programs\Python\PythonXX\Lib\idlelib\idle.bat`).
2. **Run the Script**:
   - Right-click `csv_to_excel.py` in File Explorer.
   - Select "Open with" > "IDLE" (or "Edit with IDLE" if set as default).
   - In IDLE, press `F5` or select "Run" > "Run Module" to execute the script.
   - File picker dialogs will appear for input and output file selection.
3. **Note**: If a console window appears briefly, this is normal unless compiled with `--noconsole` (see below). Ensure `xlsxwriter` is installed, or the script will fail with an error in IDLE.

### Script Behavior
- The script automatically detects the delimiter (comma or tab).
- It converts the input file to an Excel file with text-formatted cells and auto-sized columns, preserving Unicode characters.
- If the output file is open (e.g., in Excel), it prompts you to retry, save with a new name, or cancel.
- No additional GUI elements or advanced features are included, keeping the tool lightweight and straightforward.

## Compiling to an Executable (.exe)
To create a standalone .exe file (Windows only) without a console window, use `PyInstaller`.

### Steps to Compile
1. **Install PyInstaller**:
   ```bash
   pip install pyinstaller
   ```
2. **Compile the Script**:
   - In the terminal, navigate to the directory containing `csv_to_excel.py`.
   - Run:
     ```bash
     pyinstaller --onefile --noconsole csv_to_excel.py
     ```
   - The `--onefile` flag creates a single .exe file, and `--noconsole` ensures no command prompt window appears when running the executable.
3. **Locate the Executable**:
   - After compilation, find the .exe in the `dist` folder (e.g., `dist/csv_to_excel.exe`).
4. **Run the Executable**:
   - Double-click the .exe. It works the same as the Python script but doesn’t require Python to be installed on the target machine, and no console window will appear.

### Notes on Compilation
- The .exe may be large (~10-20 MB) due to bundled Python dependencies.
- Antivirus software might flag the .exe; this is common with PyInstaller. Add an exception or sign the executable if needed.
- For a smaller .exe, ensure your Python environment is clean (use a virtual environment).
- The `--noconsole` option is ideal for scripts with minimal GUI elements like this one, as it suppresses the command prompt window.

## Troubleshooting
- **Module Not Found**: Ensure `xlsxwriter` is installed in the correct Python environment (`pip show xlsxwriter`). If running in IDLE, verify the Python version matches the one where `xlsxwriter` is installed.
- **File Locked Error**: Close the output file in Excel before running the script or choose a new filename when prompted.
- **Delimiter Issues**: The script auto-detects commas or tabs. For other delimiters, modify the `detect_delimiter` function in the script.
- **Unicode Issues**: The script uses UTF-8 encoding to handle Unicode characters. If issues arise, ensure the input file is UTF-8 encoded.
- **Right-Click Issues (Windows)**: If right-clicking doesn’t show "Edit with IDLE," verify IDLE is installed and associated with `.py` files (see above). If errors occur in IDLE, ensure `xlsxwriter` is installed for the Python version used by IDLE.

## License
This project is licensed under the MIT License. Feel free to use and modify it as needed.
