import pyodbc
import pandas as pd
import json

def analyze_mdb_file(file_path):
    """Analyze the structure of an Access MDB file"""
    # Connection string for Access database
    conn_str = (
        r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        r'DBQ=' + file_path + ';'
    )
    
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        
        # Get all tables
        tables = []
        for row in cursor.tables(tableType='TABLE'):
            if not row.table_name.startswith('MSys'):  # Skip system tables
                tables.append(row.table_name)
        
        print(f"Tables found in database: {tables}")
        
        # Analyze tblSchedule structure
        if 'tblSchedule' in tables:
            print("\nAnalyzing tblSchedule structure:")
            
            # Get column information
            columns = []
            for row in cursor.columns(table='tblSchedule'):
                columns.append({
                    'name': row.column_name,
                    'type': row.type_name,
                    'size': row.column_size,
                    'nullable': row.nullable
                })
            
            print(f"Columns in tblSchedule:")
            for col in columns:
                print(f"  - {col['name']}: {col['type']} (size: {col['size']}, nullable: {col['nullable']})")
            
            # Get sample data
            cursor.execute("SELECT TOP 5 * FROM tblSchedule")
            rows = cursor.fetchall()
            
            if rows:
                print("\nSample data from tblSchedule:")
                df = pd.DataFrame.from_records(rows, columns=[col['name'] for col in columns])
                print(df.to_string())
            else:
                print("\nNo data found in tblSchedule")
        
        conn.close()
        return tables, columns if 'tblSchedule' in tables else []
        
    except Exception as e:
        print(f"Error analyzing database: {e}")
        return [], []

def analyze_xlsx_file(file_path):
    """Analyze the structure of an Excel file"""
    try:
        # Read Excel file
        xl_file = pd.ExcelFile(file_path)
        
        print(f"\nSheets found in Excel file: {xl_file.sheet_names}")
        
        for sheet_name in xl_file.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            print(f"\nSheet: {sheet_name}")
            print(f"Columns: {list(df.columns)}")
            print(f"Shape: {df.shape}")
            
            if not df.empty:
                print(f"\nSample data from {sheet_name}:")
                print(df.head().to_string())
        
        return xl_file.sheet_names
        
    except Exception as e:
        print(f"Error analyzing Excel file: {e}")
        return []

if __name__ == "__main__":
    import os
    
    print("=== Analyzing .tw2 (MDB) file ===")
    mdb_file = r"S:\Projects\936290 - UND Flight Operations Building - Grand Forks\Submittals\Titus VAVs\Data\936290 - UND Flight Operations.tw2"
    tables, columns = analyze_mdb_file(mdb_file)
    
    print("\n=== Analyzing XLSX file ===")
    xlsx_file = r"C:\Users\Jacob\Claude\VAV\~936290 Sales Plans BP4 thru addm 2 - HJA.xlsx"
    sheets = analyze_xlsx_file(xlsx_file)