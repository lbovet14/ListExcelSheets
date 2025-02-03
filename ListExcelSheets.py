import os
from openpyxl import load_workbook

def list_excel_sheets_in_directory(directory):
    """
    List all sheet names in all Excel files within the given directory and its subdirectories.

    Args:
        directory (str): The path of the directory to search.

    Returns:
        dict: A dictionary where keys are file paths and values are lists of sheet names.
    """
    excel_sheets = {}

    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(('.xlsx', '.xlsm')):  # Check for Excel file extensions
                file_path = os.path.join(root, file)
                try:
                    workbook = load_workbook(filename=file_path, read_only=True)
                    sheet_names = workbook.sheetnames
                    excel_sheets[file_path] = sheet_names
                except Exception as e:
                    print(f"Error reading {file_path}: {e}")

    return excel_sheets

def main():
    directory = os.getcwd()  # Use the current working directory
    excel_sheets = list_excel_sheets_in_directory(directory)

    if excel_sheets:
        print("\nExcel Sheets Found:")
        for file_path, sheets in excel_sheets.items():
            print(f"\nFile: {file_path}")
            for sheet in sheets:
                print(f"  - {sheet}")
    else:
        print("No Excel sheets found in the specified directory.")

if __name__ == "__main__":
    main()
