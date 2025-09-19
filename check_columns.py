import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from app import read_tw2_data_safe

def check_tw2_columns():
    """Check the columns in the specific TW2 file"""
    file_path = r"S:\Projects\936290 - UND Flight Operations Building - Grand Forks\Submittals\Titus VAVs\Data\936290 - UND Flight Operations.tw2"
    
    print(f"Checking TW2 file: {file_path}")
    print("=" * 60)
    
    try:
        result = read_tw2_data_safe(file_path)
        
        if result['success']:
            print("Successfully read TW2 file")
            print(f"Total records: {result['row_count']}")
            print(f"Total columns: {len(result['columns'])}")
            print("\nAll columns in tblSchedule:")
            print("-" * 40)
            
            # Look specifically for HeatingPrime and related fields
            heating_fields = []
            cfm_fields = []
            
            for i, column in enumerate(result['columns'], 1):
                print(f"{i:2d}. {column}")
                
                # Check for heating-related fields
                if 'heating' in column.lower() or 'heat' in column.lower():
                    heating_fields.append(column)
                    
                # Check for CFM-related fields  
                if 'cfm' in column.lower():
                    cfm_fields.append(column)
            
            print(f"\nHeating-related fields found:")
            for field in heating_fields:
                print(f"   - {field}")
                
            print(f"\nCFM-related fields found:")  
            for field in cfm_fields:
                print(f"   - {field}")
                
            # Check specifically for the fields we're looking for
            target_fields = ['HeatingPrime', 'HeatingPrimaryAirflow', 'CFMMin', 'CFMMinPrime', 'HWCFM']
            print(f"\nChecking for specific target fields:")
            for field in target_fields:
                if field in result['columns']:
                    print(f"   [FOUND] {field}")
                else:
                    print(f"   [NOT FOUND] {field}")
            
        else:
            print("Failed to read TW2 file")
            print(f"Error: {result['error']}")
            
    except Exception as e:
        print(f"Exception occurred: {str(e)}")

if __name__ == "__main__":
    check_tw2_columns()