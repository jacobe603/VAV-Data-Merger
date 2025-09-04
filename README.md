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

## Quick Start

1. **Install Dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Start the Application**:
   ```bash
   python app.py
   ```

3. **Open in Browser**:
   Navigate to `http://127.0.0.1:5004`

4. **Click Instructions**: Use the built-in instructions button for complete workflow guidance

## System Requirements

- Python 3.x
- Windows (for Microsoft Access ODBC Driver)
- Microsoft Access ODBC Driver
- Web browser

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
VAV/
├── app.py                           # Main Flask application
├── requirements.txt                 # Python dependencies
├── VAV_Data_Merger_Instructions.md  # Complete workflow guide
├── templates/
│   └── index.html                  # Web interface
├── static/
│   ├── css/
│   │   └── style.css              # Custom styling
│   └── VAV_Data_Merger_Instructions.md
├── analyze_db.py                   # Database analysis utility
├── simple_test.py                  # Basic TW2 file test
└── test_new_tw2.py                # TW2 validation test
```

## Technical Details

### Key Fixes Implemented
- **UTF-16 Encoding**: Bypasses `cursor.columns()` method that causes encoding errors
- **JSON Serialization**: Handles NaN values and various data types safely  
- **SQL Parameter Limits**: Uses batched updates to avoid Access ODBC parameter restrictions
- **Tag Normalization**: Converts between different tag naming conventions

### Known Issues
- Some SQL batch operations may encounter parameter errors (Batch 2) but core functionality remains intact
- Partial updates still succeed with 5/7 fields being updated successfully

## License

This project is developed for internal use with Titus Teams software integration.

## Support

For technical issues or questions about the workflow, refer to the built-in instructions accessible via the web interface.
