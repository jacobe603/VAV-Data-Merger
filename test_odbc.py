import pyodbc
import os

# List available ODBC drivers
print("Available ODBC Drivers:")
print("-" * 50)
drivers = pyodbc.drivers()
for driver in drivers:
    print(f"  - {driver}")

print("\n" + "=" * 50 + "\n")

# Try to connect to the TW2 file
tw2_file = r"C:\Users\Jacob\Claude\VAV\936290 - UND Flight Operations.tw2"

if os.path.exists(tw2_file):
    print(f"TW2 file found: {tw2_file}")
    print(f"File size: {os.path.getsize(tw2_file)} bytes")
    
    # Try different connection methods
    print("\nAttempting to connect to database...")
    
    # Method 1: Direct connection with .tw2 extension
    try:
        conn_str = (
            r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            f'DBQ={tw2_file};'
        )
        conn = pyodbc.connect(conn_str)
        print("[OK] Successfully connected using .tw2 extension directly")
        
        cursor = conn.cursor()
        tables = [row.table_name for row in cursor.tables(tableType='TABLE') 
                 if not row.table_name.startswith('MSys')]
        print(f"[OK] Found {len(tables)} tables")
        
        if 'tblSchedule' in tables:
            cursor.execute("SELECT COUNT(*) FROM tblSchedule")
            count = cursor.fetchone()[0]
            print(f"[OK] tblSchedule contains {count} records")
        
        conn.close()
        
    except Exception as e:
        print(f"[FAIL] Direct .tw2 connection failed: {e}")
        
        # Method 2: Copy to .mdb and try
        print("\nTrying with .mdb copy...")
        import shutil
        
        mdb_file = tw2_file.replace('.tw2', '_copy.mdb')
        shutil.copy2(tw2_file, mdb_file)
        
        try:
            conn_str = (
                r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
                f'DBQ={mdb_file};'
            )
            conn = pyodbc.connect(conn_str)
            print("[OK] Successfully connected using .mdb copy")
            
            cursor = conn.cursor()
            tables = [row.table_name for row in cursor.tables(tableType='TABLE') 
                     if not row.table_name.startswith('MSys')]
            print(f"[OK] Found {len(tables)} tables")
            
            if 'tblSchedule' in tables:
                cursor.execute("SELECT COUNT(*) FROM tblSchedule")
                count = cursor.fetchone()[0]
                print(f"[OK] tblSchedule contains {count} records")
            
            conn.close()
            
            # Clean up
            os.remove(mdb_file)
            
        except Exception as e2:
            print(f"[FAIL] .mdb copy connection failed: {e2}")
            if os.path.exists(mdb_file):
                os.remove(mdb_file)
            
else:
    print(f"TW2 file not found: {tw2_file}")