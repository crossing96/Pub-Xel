import os
import re
import shutil
# import send2trash
import tkinter as tk
from tkinter import messagebox


def center_window(rootf):
    rootf.update_idletasks()  # Ensure that window has been updated
    screen_width = rootf.winfo_screenwidth()
    screen_height = rootf.winfo_screenheight()
    window_width = rootf.winfo_reqwidth()
    window_height = rootf.winfo_reqheight()
    x = (screen_width - window_width) // 2
    y = (screen_height - window_height) // 2
    rootf.geometry(f"+{x}+{y}")

def string_to_list(input):
    #tab and line break to delimiter
    input = input.replace("\t","|")
    input = input.replace("\r\n","|")
    # 중복되는 | 제거
    to_remove = "|"
    pattern = "(?P<char>[" + re.escape(to_remove) + "])(?P=char)+"
    input = re.sub(pattern, r"\1", input)
    input = str.split(input,sep="|")
    # list 공백값 제거
    while "" in input: 
        input.remove("")
    return input

def process_ids(ids, directory):
    if isinstance(ids, str):
        ids = [ids]

    valid_ids = []
    pubmed_ids = []
    non_pubmed_valid_ids = []
    invalid_ids = []
    valid_ids_with_m_files = []
    valid_ids_without_m_files = []
    pubmed_ids_with_s_files = []
    pubmed_ids_without_s_files = []
    pubmed_ids_with_m_files = []
    pubmed_ids_without_m_files = []
    nonpubmed_ids_with_m_files = []
    nonpubmed_ids_without_m_files = []
    all_m_files = []
    all_s_files = []
    

    for id in ids:
        if re.match("^[0-9]+$", id):
            valid_ids.append(id)
            pubmed_ids.append(id)
            m_files = []
            s_files = []
            for filename in os.listdir(directory):
                if re.match(f"^{id}[^0-9]", filename):
                    if re.match(f"^{id}\.[^.]+$", filename):
                        m_files.append(filename)
                    elif re.match(f"^{id}[^.]+", filename) or re.match(f"^{id}\.[^.]+\..+$", filename):
                        s_files.append(filename)
            if m_files:
                valid_ids_with_m_files.append(id)
                pubmed_ids_with_m_files.append(id)
                all_m_files.extend(m_files)
            else:
                valid_ids_without_m_files.append(id)
                pubmed_ids_without_m_files.append(id)
            if s_files:
                pubmed_ids_with_s_files.append(id)
                all_s_files.extend(s_files)
            else:
                pubmed_ids_without_s_files.append(id)
        elif re.match("^[^0-9].*$", id):
            valid_ids.append(id)
            non_pubmed_valid_ids.append(id)
            m_files = []
            for filename in os.listdir(directory):
                if filename.startswith(id) and re.match(f"^{id}\.[^.]+$", filename):
                    m_files.append(filename)
            if m_files:
                valid_ids_with_m_files.append(id)
                nonpubmed_ids_with_m_files.append(id)
                all_m_files.extend(m_files)
            else:
                valid_ids_without_m_files.append(id)
                nonpubmed_ids_without_m_files.append(id)
        else:
            invalid_ids.append(id)

    # 0 valid_ids, 1 pubmed_ids, 2 non_pubmed_valid_ids, 3 invalid_ids, 
    # 4 valid_ids_with_m_files, 5 valid_ids_without_m_files, 
    # 6 pubmed_ids_with_m_files, 7 pubmed_ids_without_m_files, 
    # 8 pubmed_ids_with_s_files, 9 pubmed_ids_without_s_files, 
    # 10 all_m_files, 11 all_s_files
    # 12 nonpubmed_ids_with_m_files, 13 nonpubmed_ids_without_m_files
    return valid_ids, pubmed_ids, non_pubmed_valid_ids, invalid_ids, valid_ids_with_m_files, valid_ids_without_m_files, pubmed_ids_with_m_files, pubmed_ids_without_m_files, pubmed_ids_with_s_files, pubmed_ids_without_s_files, all_m_files, all_s_files,nonpubmed_ids_with_m_files,nonpubmed_ids_without_m_files

# def extract_files(requestfiles,libdir,outdir):
#     #delete destination files
#     for file_name in os.listdir(outdir):
#         send2trash.send2trash(os.path.join(outdir,file_name))

#     #copy wanted files
#     src_files = os.listdir(libdir)
#     for file_name in list(set(requestfiles) & set(src_files)):
#         full_file_name = os.path.join(libdir, file_name)
#         if os.path.isfile(full_file_name):
#             shutil.copy(full_file_name, outdir)

#     #open folder
#     os.startfile(os.path.realpath(outdir))

def extract_files_nodelete(requestfiles, libdir, outdir):
    root = tk.Tk()
    root.withdraw()

    # Check if "requestfiles" list is empty
    if not requestfiles:
        messagebox.showwarning("Warning", "The list of requested files is empty.")
        return

    # Check if there are any files in "outdir"
    if os.listdir(outdir):
        # Create a custom dialog
        dialog = tk.Toplevel(root)
        label = tk.Label(dialog, text="The output folder is not empty. Please empty the output folder and then press \"proceed\"")
        label.pack(padx=10, pady=10)

        # Create a frame for the buttons
        button_frame = tk.Frame(dialog)
        button_frame.pack(padx=10, pady=10)
        center_window(dialog)

        # Button 1: Open output folder
        open_button = tk.Button(button_frame, text="Open output folder", command=lambda: os.startfile(os.path.realpath(outdir)))
        open_button.pack(side=tk.LEFT, padx=5)

        # Button 2: Proceed
        proceed_button = tk.Button(button_frame, text="Proceed", command=lambda: proceed(dialog, outdir, requestfiles, libdir))
        proceed_button.pack(side=tk.LEFT, padx=5)

        # Button 3: Cancel
        cancel_button = tk.Button(button_frame, text="Cancel", command=dialog.destroy)
        cancel_button.pack(side=tk.LEFT, padx=5)

        root.wait_window(dialog)

        # If the dialog was destroyed (cancelled), return without copying files
        if not dialog.winfo_exists():
            return
    else:
        copy_files_and_show_success_message(outdir, requestfiles, libdir)

def proceed(dialog, outdir, requestfiles, libdir):
    if os.listdir(outdir):
        messagebox.showwarning("Warning", "The output directory is still not empty.")
    else:
        dialog.destroy()
        copy_files_and_show_success_message(outdir, requestfiles, libdir)

def copy_files_and_show_success_message(outdir, requestfiles, libdir):
    # Copy wanted files
    src_files = os.listdir(libdir)
    copied_files = 0
    for file_name in list(set(requestfiles) & set(src_files)):
        full_file_name = os.path.join(libdir, file_name)
        if os.path.isfile(full_file_name):
            shutil.copy(full_file_name, outdir)
            copied_files += 1

    # Open folder
    os.startfile(os.path.realpath(outdir))

    # Show a message box with the number of copied files
    messagebox.showinfo("Info", f"A total of {copied_files} file(s) were successfully copied.")

