import tkinter as tk
from tkinter.filedialog import askopenfilenames
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import askdirectory
import win32com.client as win_client
import os

def menu():
    while(True):
        print('Select an Option:')
        print('\t1: Convert Microsoft Word files to .pdfs')
        print('\t2: Convert all sheets in a Microsoft Excel file to .pdfs')
        print('\t0: Exit')

        match(input('Option: ')):
            case '0':
                return
            case '1':
                convert_word()
            case '2':
                convert_excel()
            case _:
                print("Error: Invalid option selected.")

def get_save_dir() -> str | None:
    save_dir = askdirectory(initialdir = os.path.realpath(__file__), mustexist = True, title='Select directory to save converted files to')
    if save_dir == '':
        save_dir = None
    return save_dir

def convert_word() -> None:
    save_dir = get_save_dir();
    files = askopenfilenames(title="Select files to convert")
    word = win_client.Dispatch("Word.application")
    word.Visible = False
    for file in files:
        file = os.path.abspath(file)
        if not os.path.exists(file):
            print("Cannot find: '" + file + "'")
            continue

        save_file =  os.path.abspath(save_dir + "/" + file.split(sep="\\")[-1] + ".pdf")
        print("Saving '" + file + "' to '" +save_file + "'")
        try:
            word_doc = word.Documents.Open(file)
            word_doc.SaveAs2(save_file, FileFormat=17)
        except:
                print("Error when saving '" + file + "' to '" +save_file + "'")
        finally:
            try:
                word_doc.Close()
            except:
                pass
    word.Quit()
    pass

def convert_excel() -> None:
    save_dir = get_save_dir();

    excel = win_client.Dispatch("Excel.application")
    file = askopenfilename(title="Select file to convert")
    
    file = os.path.abspath(file)
    if not os.path.exists(file):
        print("Cannot find: '" + file + "'")
        return

    book = excel.Workbooks.Open(file)
    for sheet in book.Worksheets:
        try:
            if(sheet.Name == "Marksheet" or sheet.Name == "Namesheet"):
                continue

            print("Saving '" + file + "'s' " + sheet.Name + " to '" + save_dir + "/" + sheet.Name + "'")
            save_file =  os.path.abspath(save_dir + "/" + sheet.Name + ".pdf")
            sheet.ExportAsFixedFormat(0, save_file)
        except:
            print("Error when saving '" + file + "'s' " + sheet.Name + " to '" + save_dir + "/" + sheet.Name + "'")
    book.Close()
    excel.Quit()


if __name__ == "__main__":
    tk.Tk().withdraw()
    menu()