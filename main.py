import credentials
import os
import win32com.client
import time


def get_most_recent_file(file_name):
    documents_folder = os.path.expanduser('C://Users') + '/Documents/'  # Assuming user's documents folder is used

    # Get a list of files with the same name
    files = [f for f in os.listdir(documents_folder) if f.startswith(file_name)]

    if not files:
        return None

    # Get the most recent file based on modification time
    most_recent_file = max(files, key=lambda f: os.path.getmtime(os.path.join(documents_folder, f)))
    update_time = f"_{time.strftime('%Y%m%d-%H%M%S')}.xlsx"
    print(os.path.join(documents_folder, most_recent_file))
    return os.path.join(documents_folder, most_recent_file)


def open_protected_excel(file_path, password):
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb = excel.Workbooks.Open(file_path, False, True, None, password)
        return wb
    except Exception as e:
        print("Failed to open the Excel file:", e)
        return None


def main():
    # Specify the file name
    file_name = 'Copy of Export_RenewalCenter.xlsx'

    # Get the most recent file with the specified name
    file_path = get_most_recent_file(file_name)
    if not file_path:
        print("No file found with the specified name.")
        return

    # Specify the password
    password = credentials.password

    # Open the Excel file
    workbook = open_protected_excel(file_path, password)
    if workbook:
        print("Excel file opened successfully.")
        # Do further processing here if needed
        # Remember to close the workbook and quit Excel after use
        workbook.Close(False)
    else:
        print("Failed to open the Excel file.")


if __name__ == "__main__":
    main()
