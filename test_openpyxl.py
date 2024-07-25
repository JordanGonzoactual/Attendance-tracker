try:
    import openpyxl
    print("openpyxl is installed and import successfully")
except ImportError as e:
    print(f"Error importing openpyxl: {e}")