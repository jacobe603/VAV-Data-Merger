"""Utility functions for database operations."""
import os
import pyodbc


def get_mdb_connection(file_path):
    """Create a connection to the Access database with error handling.

    Args:
        file_path: Path to the .tw2 or .mdb database file

    Returns:
        pyodbc.Connection object

    Raises:
        FileNotFoundError: If database file doesn't exist
        Exception: If connection fails with all drivers
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Database file not found: {file_path}")

    abs_path = os.path.abspath(file_path)

    # Try different connection strings
    connection_strings = [
        f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={abs_path};',
        f'DRIVER={{Microsoft Access Driver (*.mdb)}};DBQ={abs_path};',
        f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={abs_path};PWD=;'
    ]

    for conn_str in connection_strings:
        try:
            return pyodbc.connect(conn_str)
        except Exception:
            continue

    raise Exception(f"Failed to connect to database with any driver")


def get_project_name_from_tw2(file_path):
    """Query tblProjectInfo in TW2 database to get project name.

    Args:
        file_path: Path to the TW2 database file

    Returns:
        Project name string or None if not found/error
    """
    try:
        if not file_path or not os.path.exists(file_path):
            return None

        conn = get_mdb_connection(file_path)
        cursor = conn.cursor()

        # Query the project name from tblProjectInfo
        cursor.execute("SELECT [Name] FROM [tblProjectInfo]")
        result = cursor.fetchone()

        conn.close()

        if result:
            project_name = result[0]
            if project_name:
                return str(project_name).strip()

        return None

    except Exception as e:
        print(f"Error querying project name from TW2: {str(e)}")
        return None
