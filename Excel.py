try:
    import tkinter as tk
    from tkinter.ttk import Button
    from tkinter.filedialog import askopenfilename
    from tkinter.filedialog import askdirectory
    import win32com.client as win_client
    import sys
    import os
except:
    print("Error missing packages. Delete .venv folder and rerun Excel.bat")
    quit()

def button_callback():
    print("Converting")
    
    selected_sheets = []
    for index in listbox_sheets.curselection():
        selected_sheets.append(listbox_sheets.get(index))

    for excel_sheet in excel_workbook.Worksheets:
        try:         
            if excel_sheet.Name not in selected_sheets:
                continue
            
            save_file =  os.path.abspath(save_dir + "/" + excel_sheet.Name + ".pdf")
            print("Saving '" + excel_sheet.Name + "' to '" + save_file + "'")
            excel_sheet.ExportAsFixedFormat(0, save_file)

        except Exception as e:
            print("Error when saving '" + excel_sheet.Name + "' to '" + save_file + "':", e)
    print("Done!")

root = tk.Tk()
root.withdraw()

print("Select file to convert in window")
excel_file = askopenfilename(title='Select file to convert.')
if(excel_file == ""):
    print("1 - No File Selected Error: No Additional Information")
    sys.exit(1)

print("Select directory to save to in window")
save_dir = askdirectory(title='Select a directory to save .pdfs to.')
if(excel_file == ""):
    print("4 - No Directory Selected Error: No Additional Information")
    sys.exit(4)

print("Loading GUI")
try:
    excel_app = win_client.Dispatch('Excel.application')
    try:
        excel_workbook = excel_app.Workbooks.Open(excel_file)

        listbox_sheets = tk.Listbox(root, selectmode="extended")
        [listbox_sheets.insert(tk.END, x.Name) for x in excel_workbook.Worksheets]
        listbox_sheets.pack()

        button_convert = Button(root, text ="Convert to .pdf", command=button_callback)
        button_convert.pack()

        root.deiconify()
        root.mainloop()

    except Exception as e:
        print("3 - Excel Workbook Error: ", e)
        sys.exit(3)

    finally:
        excel_workbook.Close()

except Exception as e:
    print("2 - Excel Application Error: ", e)
    sys.exit(2)

finally:
    excel_app.Quit()

# Error Codes:
# 1: No file selected
# 2: Excel Application Error
# 3: Excel Workbook Error
# 4: No directory selected