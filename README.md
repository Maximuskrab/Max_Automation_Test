# Excel Appender Tool

A beautiful and professional Streamlit application that concatenates multiple Excel files from a specified folder into a single combined Excel file.

## Features

- 📁 **Folder Path Input** - Specify any folder to scan for Excel files
- 🔍 **Automatic File Discovery** - Finds all .xlsx and .xls files automatically
- ⚡ **One-Click Merge** - Concatenate all files with a single click
- ✅ **Real-time Status** - Visual feedback during processing
- 🎨 **Beautiful UI** - Professional Streamlit interface
- 🔄 **Reset Functionality** - Clear and start fresh at any time

## Installation

### Prerequisites
- Python 3.8+
- pip package manager

### Setup

1. Clone or download this repository
2. Navigate to the project directory:
   ```bash
   cd AutomationTool_1
   ```

3. Create and activate a virtual environment (recommended):
   ```bash
   # Windows
   python -m venv venv
   venv\Scripts\activate
   
   # macOS/Linux
   python3 -m venv venv
   source venv/bin/activate
   ```

4. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Start the Streamlit app:
   ```bash
   streamlit run src/app.py
   ```

2. Open your browser (usually opens automatically at `http://localhost:8501`)

3. **Enter Folder Path** - Input the full path to your folder containing Excel files
   - Example: `C:\Users\YourName\Documents\ExcelFiles`
   - Example: `/Users/YourName/Documents/ExcelFiles`

4. **Review Files** - The app will show how many Excel files it found

5. **Click Run** - Process and merge all files

6. **Download Output** - The combined Excel file is saved with a timestamp

7. **Reset** - Click Reset to clear and process another folder

## Output

- Combined Excel file is saved as: `output/appended_<timestamp>.xlsx`
- All sheets from all Excel files are concatenated into a single dataset
- Original files remain unchanged

## Project Structure

```
AutomationTool_1/
├── .streamlit/
│   └── config.toml                 # Streamlit configuration
├── src/
│   ├── app.py                      # Main Streamlit application
│   └── excel_appender/
│       ├── __init__.py
│       └── core.py                 # Excel processing logic
├── .gitignore
├── README.md
└── requirements.txt
```

## Dependencies

- **pandas** - Data manipulation and Excel file handling
- **openpyxl** - Excel file support for pandas
- **streamlit** - Web UI framework

## Troubleshooting

### Error: "No Excel files found in folder"
- Ensure the folder path is correct
- Check that the folder contains .xlsx or .xls files
- Verify you have read permissions for the folder

### Error: "Cannot read file"
- The file might be corrupted or locked by another application
- Try closing the file in Excel
- Ensure the file is a valid Excel file

### Error: "Permission denied"
- Check folder access permissions
- Run the application with appropriate permissions
- Try with a different folder
