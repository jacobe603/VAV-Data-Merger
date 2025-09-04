# VAV Data Merger - Complete Workflow Instructions

## Overview
The VAV Data Merger is a web-based tool that combines Excel spreadsheet data with Titus Teams TW2 database files to populate VAV unit specifications efficiently. This eliminates manual data entry and ensures consistency across your VAV scheduling workflow.

## System Requirements
- Windows operating system
- Python 3.x installed
- Microsoft Access ODBC Driver
- Web browser (Chrome, Firefox, Edge)
- Access to Titus Teams software

---

## Part 1: Titus Teams Preparation

### Step 1: Create New Job in Titus Teams
1. Open Titus Teams software
2. Create a new job/project
3. Set up basic project parameters:
   - Project name (e.g., "936290 - UND Flight Operations")
   - Building/zone information
   - System specifications

### Step 2: Create the Master VAV Unit
1. Navigate to the VAV/Terminal Unit section in Titus Teams
2. Create the first VAV unit with complete specifications:
   - **Tag**: Use format V-1-01 (note the zero-padded format)
   - **Unit Size**: Specify the VAV box size (e.g., 6", 8", 10", 14")
   - **Inlet Size**: Air inlet diameter
   - **CFM Design**: Maximum design airflow
   - **CFM Min Prime**: Minimum primary airflow
   - **Heating Prime**: Primary airflow during heating
   - **HWCFM**: Hot water coil airflow
   - **HWGPM**: Hot water flow rate in gallons per minute

### Step 3: Copy VAV Units to Match Quantity
1. Select your completed VAV unit
2. Use Titus Teams' copy/duplicate function
3. Create the total number of units needed (e.g., 42 units for this project)
4. Ensure tags follow the sequential format: V-1-01, V-1-02, V-1-03, etc.
5. Leave other specifications as placeholders (they will be updated by our tool)

### Step 4: Export TW2 Database
1. Save your Titus Teams project
2. Locate the TW2 database file (typically in your project folder)
3. Copy the .tw2 file to a known location (e.g., `C:\VAV_Projects\`)
4. **Important**: Keep the .tw2 extension - this is crucial for the merger tool

---

## Part 2: Excel Data Preparation

### Required Excel Format
Your Excel file must contain the following columns with exact headers:

| Column Name | Description | Example Values |
|-------------|-------------|----------------|
| Unit_No | Unit identifier matching TW2 tags | V-1-1, V-1-2, V-1-3 |
| Unit_Size | VAV box size designation | 6, 8, 10, 14, 24x16 |
| Inlet_Size | Air inlet size in inches | 6", 8", 10", 14" |
| CFM_Max | Maximum design airflow | 500, 750, 1200 |
| CFM_Min | Minimum airflow | 100, 150, 250 |
| CFM_Heat | Heating mode airflow | 300, 450, 600 |
| GPM | Hot water flow rate | 2.5, 3.0, 4.5 |

### Excel Format Requirements
1. Data should start in **row 3** (rows 1-2 can be headers/titles)
2. Remove any completely blank rows
3. Ensure Unit_No column matches your TW2 tags (V-1-1 format is acceptable)
4. Save as .xlsx format

### Tag Format Notes
- Excel typically uses: V-1-1, V-1-2, V-1-3
- TW2 uses: V-1-01, V-1-02, V-1-03
- The merger tool automatically handles this conversion

---

## Part 3: Using the VAV Data Merger Tool

### Step 1: Start the Application
1. Open Command Prompt or PowerShell
2. Navigate to the VAV tool directory:
   ```
   cd C:\Users\Jacob\Claude\VAV
   ```
3. Start the Flask application:
   ```
   python app_fixed.py
   ```
4. Wait for the message: "Running on http://127.0.0.1:5004"
5. Open your web browser and navigate to: `http://127.0.0.1:5004`

### Step 2: Upload TW2 Database File
1. In the **TW2 Database** section (left side), drag and drop your .tw2 file
2. **Expected Result**: 
   - Green success message showing number of records found
   - Data preview table showing current TW2 values
   - Sample data with columns: Tag, UnitSize, InletSize, CFMDesign, HWCFM, HWGPM

### Step 3: Upload Excel Spreadsheet
1. In the **Excel Data** section (right side), drag and drop your .xlsx file
2. **Expected Result**:
   - Green success message showing number of records processed
   - Data preview table showing Excel values
   - Sample data with columns: Unit_No, Unit_Size, Inlet_Size, CFM_Max, CFM_Min, GPM

### Step 4: Review Field Mapping
Once both files are loaded, the **Field Mapping** section will appear:

| TW2 Database Field | → | Excel Spreadsheet Field | Description |
|-------------------|---|------------------------|-------------|
| Tag | → | Unit_No | Unit identifier - Must match Excel Unit_No |
| UnitSize | → | Unit_Size | VAV unit size designation |
| InletSize | → | Inlet_Size | Air inlet size in inches |
| CFMDesign | → | CFM_Max | Design air flow rate in CFM - Maximum airflow |
| CFMMinPrime | → | CFM_Min | Minimum primary airflow in CFM |
| HeatingPrime | → | CFM_Heat | Primary airflow during heating mode in CFM |
| HWCFM | → | CFM_Heat | Hot water coil airflow in CFM |
| HWGPM | → | GPM | Hot water flow rate in GPM |

**Note**: Hover over TW2 field names to see detailed descriptions and examples.

### Step 5: Apply the Mapping
1. Review the mapping table to ensure it's correct
2. Click the **"Apply Mapping"** button
3. **Monitor Progress**: The tool will process each unit in batches
4. **Expected Results**:
   - Progress messages in the browser console
   - Final success message showing:
     - Number of records updated
     - Backup file location
     - Any errors encountered

### Step 6: Verify Results
**Success Indicators**:
- Green success message: "Mapping applied successfully!"
- "Updated X records" (should match your Excel row count)
- "Backup created: [filename].backup_YYYYMMDD_HHMMSS"

**If Errors Occur**:
- Check tag format matching between Excel and TW2
- Verify all required columns exist in Excel
- Ensure no blank rows in critical data

---

## Part 4: Post-Processing & Titus Teams Integration

### Step 1: Verify the Updated TW2 File
1. The original TW2 file has been updated with Excel data
2. A backup file has been created with timestamp
3. **Test the file** (optional):
   ```
   python test_new_tw2.py
   ```

### Step 2: Import Back to Titus Teams
1. Open your Titus Teams project
2. **Option A**: Replace the existing TW2 database file
   - Close Titus Teams completely
   - Replace the original TW2 file with your updated version
   - Reopen Titus Teams
   
2. **Option B**: Import/Refresh database (if available in your Titus version)
   - Use Teams' import/refresh function
   - Point to your updated TW2 file

### Step 3: Quality Assurance
1. Open a few VAV units in Titus Teams
2. Verify the data has been populated correctly:
   - Unit sizes match Excel specifications
   - CFM values are populated
   - GPM values are correct
3. Run any built-in Titus Teams validation checks
4. Generate a report to confirm all units have proper data

### Step 4: Project Completion
1. Save your Titus Teams project
2. Run final calculations/analysis in Teams
3. Generate required reports and outputs
4. Archive your backup files for future reference

---

## Part 5: Troubleshooting Guide

### Common Issues and Solutions

#### "Error loading TW2 file: 'utf-16-le' codec can't decode bytes"
- **Cause**: Database encoding issue
- **Solution**: This is resolved in app_fixed.py - ensure you're using the correct application version

#### "Upload error: SyntaxError: Unexpected token 'N'"
- **Cause**: NaN values in Excel data
- **Solution**: Clean up Excel file, remove empty cells, ensure numeric columns contain valid numbers

#### "Too few parameters. Expected X"
- **Cause**: SQL parameter mismatch
- **Solution**: This is resolved with batched updates in app_fixed.py

#### "No matching tags found"
- **Cause**: Tag format mismatch between Excel and TW2
- **Solutions**:
  - Ensure Excel uses format V-1-1, V-1-2, etc.
  - Ensure TW2 uses format V-1-01, V-1-02, etc.
  - The tool handles conversion automatically

#### Application won't start
- **Check**: Python installation and PATH
- **Check**: Required packages installed (flask, pandas, pyodbc)
- **Install missing packages**:
  ```
  pip install flask flask-cors pandas pyodbc
  ```

#### "No records updated"
- **Cause**: No matching tags between files
- **Solution**: 
  - Verify tag naming consistency
  - Check that Unit_No in Excel corresponds to Tag in TW2
  - Ensure Excel data starts in the correct row (row 3)

### Getting Help
- Check the console output for detailed error messages
- Review backup files if you need to restore original data
- Verify file formats and column names match requirements exactly

---

## Summary Workflow Checklist

- [ ] Create job in Titus Teams
- [ ] Set up master VAV with complete specifications  
- [ ] Copy VAV to create required quantity (42 units)
- [ ] Export TW2 database file
- [ ] Prepare Excel file with required columns
- [ ] Start VAV Data Merger application (`python app_fixed.py`)
- [ ] Upload TW2 file via drag-and-drop
- [ ] Upload Excel file via drag-and-drop
- [ ] Review field mapping table
- [ ] Apply mapping and monitor results
- [ ] Verify backup creation and success message
- [ ] Import updated TW2 back to Titus Teams
- [ ] Perform quality assurance checks
- [ ] Complete project analysis and reporting

## File Locations
- **Application**: `C:\Users\Jacob\Claude\VAV\app_fixed.py`
- **Web Interface**: `http://127.0.0.1:5004`
- **Backup Location**: Same directory as original TW2 file
- **Test Scripts**: `test_new_tw2.py` for verification

---

*Last updated: September 2025*
*Application version: app_fixed.py with UTF-16 encoding fixes and batched SQL updates*