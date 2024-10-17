import os
import regex
import win32com.client
from pywintypes import com_error

BASE_DIR = os.path.dirname(os.path.realpath(__file__))
TRIM_FROM_DIR = '_assignsubmission_file'


# Utility

# Users options
def menu():
    while True:
        print('1: Rename folders')
        print('2: Export excel sheets to .pdf')
        print('0: Exit')
        match input('What would you like to do?: '):
            case '0':
                quit()
            case '1':
                rename_folders()
            case '2':
                convert_to_PDF()

# Check if a file path exists and returns the absolute path or None if invalid
def validate_file_path(path: str) -> str:
    # Validate path
    if os.path.isfile(path):
        # Convert relative paths to absolute path - ignored if already absolute
        if os.path.isfile(BASE_DIR + '\\' + path):
            path = BASE_DIR + '\\' + path
    # Invalid Path
    else:
        path = None
    return path

# Check if a directory path exists and returns the absolute path or None if invalid
def validate_dir_path(path: str) -> str:
    # Validate path
    if os.path.isdir(path):
        # Convert relative paths to absolute path - ignored if already absolute
        if os.path.isdir(BASE_DIR + '\\' + path):
            path = BASE_DIR + '\\' + path
    # Uses base directory
    if path == '':
        return BASE_DIR
    # Invalid Path
    else:
        path = None
    return path


# Menu Options

# Rename a folder by removing the TRIM_FROM_DIR string from it if applicable
def rename_folders():
    for dir in os.listdir(BASE_DIR):
        # Ignore Venv and non directory
        if(dir == '.venv'):
            continue
        if not os.path.isdir(dir):
            continue
        
        # Remove TRIM_FROM_DIR from dir
        if TRIM_FROM_DIR in dir:
            splits = dir.replace(TRIM_FROM_DIR, '').split('_')
            new_dir = BASE_DIR + '\\' + splits[0] + ' (' + splits[1] + ')'
            # Try to keep user's changes
            if len(splits) > 2:
                for i in range(2, len(splits)):
                    new_dir += '_' + splits[i]
            os.rename(BASE_DIR + '\\' + dir, new_dir)
        # Inform user why file has not changed
        else:
            print("'" + dir + "' does not include '" + TRIM_FROM_DIR + "' so it has not been renamed.")

# Convert an Excel documents sheets to PDFs
def convert_to_PDF():

    excel_path = validate_file_path(input('Excel path: '))

    if excel_path == None:
        print('Feedback Excel path could not be found.')
        return
    print('Reading from ' + excel_path)

    dir_path = validate_dir_path(input('PDF path: '))
    
    if dir_path == None:
        print('Feedback PDF path could not be found.')
        return
    print('Saving to ' + dir_path)

    sub_dirs = os.listdir()

    # Open Workbook
    try:
        excel_app = win32com.client.Dispatch("Excel.Application")
        workbook = excel_app.Workbooks.Open(excel_path)

        for worksheet in workbook.Worksheets:
            if worksheet.Name == 'Namesheet' or worksheet.Name == 'Marksheet':
                continue
            file_name = None
            splits = worksheet.Name.split(' ')
            if len(splits) >= 2:
                student_name = (splits[0] + ' ' + splits[1]).title()
            else:
                student_name = worksheet.Name.title()
            
            # Put feedback in folders if they contain the worksheet name
            for sub_dir in sub_dirs:
                if student_name.lower() in sub_dir.lower():
                    file_name = dir_path + '\\' + sub_dir + '\\' + student_name + "_feedback.pdf"
                    break

            # Put feedback in feedback folder if no other folder found
            if file_name == None:
                file_name = dir_path + '\\' + student_name + "_feedback.pdf"
            worksheet.ExportAsFixedFormat(0, file_name)
            print('Saved to ' + file_name)
    except com_error:
        print("Issue opening '" + excel_path + "' in Microsoft Excel")
    finally:
        workbook.Close()
        excel_app.Quit()

menu()