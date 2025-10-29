# VAV Data Merger

A web-based tool for combining Excel spreadsheet data with Titus Teams TW2 database files for VAV unit scheduling.

## Changelog

- See `CHANGELOG.md` for notable changes and release notes.

## Features

- **Drag & Drop Interface**: Easy file uploads for TW2 and Excel files
- **Automatic Field Mapping**: Smart mapping between Excel columns and TW2 database fields  
- **Data Preview**: Visual confirmation of data before processing
- **Backup Creation**: Automatic backups with timestamps before making changes
- **Built-in Instructions**: Comprehensive workflow guide accessible from the web interface
- **UTF-16 Error Handling**: Resolves Microsoft Access ODBC driver encoding issues
- **Tag Format Conversion**: Automatic conversion between Excel (V-1-1) and TW2 (V-1-01) formats

## System Requirements

- **Python**: 3.9, 3.10, 3.11, 3.12, or 3.13 (tested and verified)
- **OS**: Windows (for Microsoft Access ODBC Driver)
- **Dependencies**: Microsoft Access ODBC Driver (pre-installed on most Windows systems)
- **Browser**: Any modern web browser

## Quick Start (Recommended)

### Windows - Automated Setup
1. **Double-click `setup.bat`** in the project folder
   - This will automatically create a virtual environment and install dependencies
2. **After setup completes**, follow the on-screen instructions
3. **Open your browser** to `http://127.0.0.1:5004`

### Manual Setup (All Platforms)
1. **Create a virtual environment**:
   ```bash
   python -m venv .venv
   ```

2. **Activate the virtual environment**:
   - **Windows Command Prompt**:
     ```bash
     .venv\Scripts\activate.bat
     ```
   - **Windows PowerShell**:
     ```bash
     .venv\Scripts\Activate.ps1
     ```
   - **macOS/Linux**:
     ```bash
     source .venv/bin/activate
     ```

3. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

4. **Start the application**:
   ```bash
   python app.py
   ```

5. **Open in browser**:
   Navigate to `http://127.0.0.1:5004`

## Running the Application

- **Start**: Run `python app.py` (with virtual environment activated)
- **URL**: `http://127.0.0.1:5004`
- **Stop**: Press `Ctrl + C` in the terminal

## Troubleshooting

### Python Version Issues
- **Error during `pip install`**: Ensure you have Python 3.9 or later
  - Check version: `python --version`
  - Download from https://www.python.org/
  - Make sure to check "Add Python to PATH" during installation

### ODBC Driver Issues
- **Error: "Microsoft Access Driver not found"**
  - This is typically pre-installed on Windows
  - If missing, install Microsoft Access Database Engine from the Windows installation media
  - Restart the application after installing the driver

### Port Already in Use
- **Error: "Port 5004 is already in use"**
  - Option 1: Stop the existing process using the port
  - Option 2: Edit `app.py` and change the port number in the `app.run()` line, then restart

### Virtual Environment Not Activating
- **Windows PowerShell issue**: May need to enable script execution
  - Run: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
  - Then retry: `.venv\Scripts\Activate.ps1`

### Dependency Installation Fails
- **Clear pip cache and retry**:
  ```bash
  pip install --no-cache-dir -r requirements.txt
  ```

## File Format Requirements

### TW2 Files
- Microsoft Access database with `.tw2` extension
- Must contain `tblSchedule` table with VAV unit data

### Excel Files  
- `.xlsx` format with specific column headers:
  - `Unit_No`: Unit identifier (V-1-1, V-1-2, etc.)
  - `Unit_Size`: VAV box size (6, 8, 10, 14, 24x16)
  - `Inlet_Size`: Air inlet size (6", 8", 10", 14")
  - `CFM_Max`: Maximum design airflow
  - `CFM_Min`: Minimum airflow
  - `CFM_Heat`: Heating mode airflow  
  - `GPM`: Hot water flow rate

## Field Mapping

| TW2 Database Field | Excel Column | Description |
|-------------------|--------------|-------------|
| Tag | Unit_No | Unit identifier |
| UnitSize | Unit_Size | VAV unit size |
| InletSize | Inlet_Size | Air inlet size |
| CFMDesign | CFM_Max | Design airflow |
| CFMMinPrime | CFM_Min | Minimum airflow |
| HWCFM | CFM_Heat | Hot water coil airflow |
| HWGPM | GPM | Hot water flow rate |

## Project Structure

```
VAV-Data-Merger/
├── app.py                           # Main Flask application
├── requirements.txt                 # Python dependencies
├── setup.bat                        # Windows automated setup script
├── VAV_Data_Merger_Instructions.md  # Complete workflow guide
├── templates/
│   └── index.html                  # Web interface
├── static/
│   ├── css/
│   │   └── style.css              # Custom styling
│   └── VAV_Data_Merger_Instructions.md
├── analyze_db.py                   # Database analysis utility
├── test_odbc.py                    # ODBC connection test
└── check_columns.py                # Database column inspection
```

## Technical Details

### Key Fixes Implemented
- **UTF-16 Encoding**: Bypasses `cursor.columns()` method that causes encoding errors
- **JSON Serialization**: Handles NaN values and various data types safely  
- **SQL Parameter Limits**: Uses batched updates to avoid Access ODBC parameter restrictions
- **Tag Normalization**: Converts between different tag naming conventions
- **Flexible Dependencies**: Compatible with Python 3.9 through 3.13

### Known Issues
- Some SQL batch operations may encounter parameter errors (Batch 2) but core functionality remains intact
- Partial updates still succeed with 5/7 fields being updated successfully

## Production Deployment

For production use, deploy behind a WSGI server and reverse proxy:

### Example with Waitress (Windows-friendly)
1. **Install Waitress**:
   ```bash
   pip install waitress
   ```

2. **Run with Waitress**:
   ```bash
   python -m waitress --listen=0.0.0.0:5004 app:app
   ```

3. **Firewall**: Ensure firewall allows inbound connections if serving across a network

## License

This project is developed for internal use with Titus Teams software integration.

## Support

For technical issues or questions about the workflow:
1. Check the **Troubleshooting** section above
2. Refer to the built-in instructions accessible via the web interface
3. Review the `VAV_Data_Merger_Instructions.md` file
