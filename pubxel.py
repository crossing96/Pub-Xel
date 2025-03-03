import tkinter as tk
from tkinter import ttk
import keyboard
import threading
import time
import os
import tkinter.messagebox as messagebox
import pyperclip
import webbrowser

# Flag to indicate whether an action is in progress
action_in_progress = False

# Create a Tkinter window
root = tk.Tk()
root.title("Pub-Xel [BETA]")

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
# Build the full path to the file
file_path = os.path.join(script_dir, 'logo64.ico')
root.iconbitmap(file_path)

#lib_directory
# lib_directory = "C:\\Users\\PC\\Desktop\\research\\files\\"
# output_directory = "C:\\Users\\PC\\Desktop\\pythonOutput\\"

desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') 
output_directory = os.path.join(desktop,"article_output")
lib_directory = os.path.join(desktop,"article_library")

if not os.path.exists(lib_directory):
    os.makedirs(lib_directory)

if not os.path.exists(output_directory):
    os.makedirs(output_directory)

def force_open_folder(dir):
    if os.path.exists(dir):
        os.startfile(dir)
    else:
        os.makedirs(dir)
        time.sleep(0.3)
        os.startfile(dir)

def force_open_lib():
    force_open_folder(lib_directory)

def force_open_output():
    force_open_folder(output_directory)

def callback(result):
    # Handle the result in your higher-level code
    for filename in result:
        file_path = os.path.join(lib_directory, filename)
        os.startfile(file_path)
    print("Result:", result)

    global action_in_progress
    action_in_progress = False
    global new_window
    new_window.destroy()

    # You can perform any actions with the result here

from mainfunctions import string_to_list
from mainfunctions import process_ids

def openfile_from_list(filelist, dir):
    # If there are more than 5 files, show a warning and ask for confirmation
    if len(filelist) > 5:
        proceed = messagebox.askyesno("Warning", f"You are about to open {len(filelist)} files. Do you want to proceed?")
        if not proceed:
            return

    failed_files = []  # List to keep track of files that could not be opened
    for filename in filelist:
        try:
            filepath = os.path.join(dir, filename)
            os.startfile(filepath)
        except Exception as e:
            failed_files.append(filename)  # Add the failed file to the list
            continue

    # If there were any failed files, show a messagebox after all file opening attempts are completed
    if failed_files:
        messagebox.showerror("Error", f"Could not open the following files: {', '.join(failed_files)}")

def a_openfile():
    global action_in_progress
    action_in_progress = True  # Set the action in progress
    clipboardstring = pyperclip.paste()
    idlist = string_to_list(clipboardstring)
    output = process_ids(idlist,lib_directory)
    openfile_from_list(output[10],lib_directory)
    action_in_progress = False


def a_openadv():

    global action_in_progress
    action_in_progress = True

    def reset_action_flag():
        global action_in_progress
        action_in_progress = False
        new_window.destroy()
    
    from mainfunctions import center_window
    from mainfunctions import string_to_list
    from mainfunctions import process_ids
    from mainfunctions import extract_files_nodelete

    # Get the selected cell value from clipboard
    selected_cell_value = pyperclip.paste()
    selected_cell_value = string_to_list(selected_cell_value)
    output = process_ids(selected_cell_value,lib_directory)
    valid_ids = output[0]
    pubmed_ids = output[1]
    non_pubmed_valid_ids = output[2]
    invalid_ids = output[3]
    valid_ids_with_m_files = output[4]
    valid_ids_without_m_files = output[5]
    pubmed_ids_with_m_files = output[6]
    pubmed_ids_without_m_files = output[7]
    pubmed_ids_with_s_files = output[8]
    pubmed_ids_without_s_files = output[9]
    all_m_files = output[10]
    all_s_files = output[11]
    nonpubmed_ids_with_m_files = output[12]
    nonpubmed_ids_without_m_files = output[13]

    def copy_list(lst):
        # Check if the list is empty
        if not lst:
            return None  # Don't do anything
        # Concatenate the list with a line break as a delimiter
        result = '\n'.join(lst)
        # Copy the string to the clipboard
        pyperclip.copy(result)
    
    # Initialize the tkinter GUI
    global new_window
    new_window = tk.Toplevel(root)
    new_window.withdraw()  # Hide the window
    new_window.title("Inspect files")

    # Top frame (frame 1)
    frame1 = tk.Frame(new_window)
    frame1.pack(side='top', fill='both', expand=True, padx=5, pady=5)

    # Boolean values for each row
    row0_bool = len(pubmed_ids)>0
    row1_bool = len(pubmed_ids_without_m_files)>0
    row2_bool = len(non_pubmed_valid_ids)>0
    row3_bool = len(nonpubmed_ids_without_m_files)>0
    row4_bool = len(invalid_ids)>0

    # Create rows in frame 1
    if row0_bool:
        n1 = len(pubmed_ids)
        tk.Label(frame1, text=f"{n1} PubMed article(s):", anchor='w').grid(row=0, column=0, sticky='w', padx=5, pady=5)
        def copy_list0():
            copy_list(pubmed_ids)
        copy0=tk.Button(frame1, text="Copy", command=copy_list0)
        copy0.grid(row=0, column=2, sticky='e', padx=5, pady=5)
        def openlink0():
            webbrowser.open("https://pubmed.ncbi.nlm.nih.gov/?term="+"%5Buid%5D+OR+".join(pubmed_ids)+"%5Buid%5D&sort=date")
        search0=tk.Button(frame1, text="Search PubMed",command=openlink0)
        search0.grid(row=0, column=3, sticky='e', padx=5, pady=5)

    if row0_bool:
        n1 = len(pubmed_ids)
        n2 = len(pubmed_ids_without_m_files)
        if row1_bool:
            tk.Label(frame1, text=f"    {n2} of the {n1} {'has' if n2 == 1 else 'have'} no main file(s)", anchor='w').grid(row=1, column=0, sticky='w', padx=5, pady=5)
            
            def copy_list1():
                copy_list(pubmed_ids_without_m_files)
            copy1=tk.Button(frame1, text="Copy",command=copy_list1)
            copy1.grid(row=1, column=2, sticky='e', padx=5, pady=5)

            def openlink1():
                webbrowser.open("https://pubmed.ncbi.nlm.nih.gov/?term="+"%5Buid%5D+OR+".join(pubmed_ids_without_m_files)+"%5Buid%5D&sort=date")

            search1=tk.Button(frame1, text="Search PubMed",command=openlink1)
            search1.grid(row=1, column=3, sticky='e', padx=5, pady=5)
        else:
            tk.Label(frame1, text="    Main files all available", anchor='w').grid(row=1, column=0, sticky='w', padx=5, pady=5)

    if row2_bool:
        n1 = len(non_pubmed_valid_ids)
        tk.Label(frame1, text=f"{n1} Non-PubMed article(s):", anchor='w').grid(row=2, column=0, sticky='w', padx=5, pady=5)
        def copy_list2():
            copy_list(non_pubmed_valid_ids)
        copy2=tk.Button(frame1, text="Copy",command=copy_list2)
        copy2.grid(row=2, column=2, sticky='e', padx=5, pady=5)

    if row2_bool:
        n1 = len(non_pubmed_valid_ids)
        n2 = len(nonpubmed_ids_without_m_files)
        if row3_bool:
            tk.Label(frame1, text=f"    {n2} of the {n1} {'has' if n2 == 1 else 'have'} no main file(s)", anchor='w').grid(row=3, column=0, sticky='w', padx=5, pady=5)
            def copy_list3():
                copy_list(nonpubmed_ids_without_m_files)
            copy3=tk.Button(frame1, text="Copy",command=copy_list3)
            copy3.grid(row=3, column=2, sticky='e', padx=5, pady=5)
        else:
            tk.Label(frame1, text="    Main files all available", anchor='w').grid(row=3, column=0, sticky='w', padx=5, pady=5)

    if row4_bool:
        n1 = len(invalid_ids)

        tk.Label(frame1, text=f"{n1} invalid ID(s):", anchor='w').grid(row=4, column=0, sticky='w', padx=5, pady=5)
        def copy_list4():
            copy_list(invalid_ids)
        copy4=tk.Button(frame1, text="Copy",command=copy_list4)
        copy4.grid(row=4, column=2, sticky='e', padx=5, pady=5)

    frame1.columnconfigure(1, weight=1)

    # Middle frame (frame 2)
    frame2 = tk.Frame(new_window)
    frame2.pack(side='top', fill='both', expand=True, padx=5, pady=5)

    # Left subframe in frame 2
    left_frame = tk.Frame(frame2)
    left_frame.pack(side='left', fill='both', expand=True, padx=5, pady=5)

    left_label_frame = tk.Frame(left_frame)
    left_label_frame.pack(side='top', fill='x')
    tk.Label(left_label_frame, text="Main files").pack()

    left_checkbox_frame = tk.Frame(left_frame)
    left_checkbox_frame.pack(side='top', fill='both', expand=True)


    canvas_main = tk.Canvas(left_checkbox_frame)
    canvas_main.pack(side="left", fill="both", expand=True)
    scrollbar_main = ttk.Scrollbar(left_checkbox_frame, orient="vertical", command=canvas_main.yview)
    scrollbar_main.pack(side="right", fill="y")
    canvas_main.configure(yscrollcommand=scrollbar_main.set)
    # Create a frame inside the canvas_main to hold the checkboxes
    checkbox_frame_main = ttk.Frame(canvas_main)
    canvas_main.create_window((0, 0), window=checkbox_frame_main, anchor="nw")
    # Create checkboxes for file selection
    main_files_checkboxes = {}
    for filename in all_m_files: 
        main_files_checkboxes[filename] = tk.IntVar()
        checkbox_main = tk.Checkbutton(checkbox_frame_main, text=f"File {filename}", variable=main_files_checkboxes[filename])
        checkbox_main.deselect()
        checkbox_main.pack(anchor="w")
    # Update the canvas_main scrolling region
    def on_canvas_main_configure(event):
        canvas_main.configure(scrollregion=canvas_main.bbox("all"))
        canvas_main.bind("<Configure>", on_canvas_main_configure)
    # Bind the MouseWheel event to the canvas_main for scrolling
    def on_mousewheel_main(event):
        canvas_main.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas_main.bind("<MouseWheel>", on_mousewheel_main)


    # Right subframe in frame 2
    right_frame = tk.Frame(frame2)
    right_frame.pack(side='right', fill='both', expand=True, padx=5, pady=5)

    right_label_frame = tk.Frame(right_frame)
    right_label_frame.pack(side='top', fill='x')
    tk.Label(right_label_frame, text="Suppl files").pack()

    right_checkbox_frame = tk.Frame(right_frame)
    right_checkbox_frame.pack(side='top', fill='both', expand=True)


    canvas_suppl = tk.Canvas(right_checkbox_frame)
    canvas_suppl.pack(side="left", fill="both", expand=True)
    scrollbar_suppl = ttk.Scrollbar(right_checkbox_frame, orient="vertical", command=canvas_suppl.yview)
    scrollbar_suppl.pack(side="right", fill="y")
    canvas_suppl.configure(yscrollcommand=scrollbar_suppl.set)
    # Create a frame inside the canvas_suppl to hold the checkboxes
    checkbox_frame_suppl = ttk.Frame(canvas_suppl)
    canvas_suppl.create_window((0, 0), window=checkbox_frame_suppl, anchor="nw")
    # Create checkboxes for file selection
    supplementary_files_checkboxes = {}
    for filename in all_s_files: 
        supplementary_files_checkboxes[filename] = tk.IntVar()
        checkbox_suppl = tk.Checkbutton(checkbox_frame_suppl, text=f"File {filename}", variable=supplementary_files_checkboxes[filename])
        checkbox_suppl.deselect()
        checkbox_suppl.pack(anchor="w")
    # Update the canvas_suppl scrolling region
    def on_canvas_suppl_configure(event):
        canvas_suppl.configure(scrollregion=canvas_suppl.bbox("all"))
    canvas_suppl.bind("<Configure>", on_canvas_suppl_configure)
    # Bind the MouseWheel event to the canvas_suppl for scrolling
    def on_mousewheel_suppl(event):
        canvas_suppl.yview_scroll(int(-1 * (event.delta / 120)), "units")
    canvas_suppl.bind("<MouseWheel>", on_mousewheel_suppl)

    # Bottom frame (frame 3)
    frame3 = tk.Frame(new_window)
    frame3.pack(side='top', fill='both', expand=True, padx=5, pady=5)

    # Create rows in frame 3
    tk.Label(frame3, text="Open:").grid(row=0, column=1, sticky='e', padx=5, pady=5)

    def open_files(files, directory):
        try:
            if not files:
                raise Exception("No files to open.")

            # Check if there are 6 or more files to open
            if len(files) >= 6:
                confirm = messagebox.askyesno("Confirm Open", "You are about to open " + str(len(files)) + " files. Proceed?")
                if not confirm:
                    return
            return(files)
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def open_selected_files():
        selected_files = [filename for filename, var in main_files_checkboxes.items() if var.get() == 1] + [filename for filename, var in supplementary_files_checkboxes.items() if var.get() == 1]
        print(selected_files)
        files = open_files(selected_files, lib_directory)
        return(files)

    def open_all_files():
        all_files = all_m_files + all_s_files
        files = open_files(all_files, lib_directory)
        return(files)

    open_selected_button = tk.Button(frame3, text="Selected", command=lambda: callback(open_selected_files()))
    open_selected_button.grid(row=0, column=2, sticky='e', padx=5, pady=5)
    open_selected_button.config(underline=0)
    new_window.bind("S", lambda _: open_selected_button.invoke())
    new_window.bind("s", lambda _: open_selected_button.invoke())

    open_all_button = tk.Button(frame3, text="All", command=lambda: callback(open_all_files()))
    open_all_button.grid(row=0, column=3, sticky='e', padx=5, pady=5)
    open_all_button.config(underline=0)   
    new_window.bind("A", lambda _: open_all_button.invoke())
    new_window.bind("a", lambda _: open_all_button.invoke())


    tk.Label(frame3, text="Export:").grid(row=1, column=1, sticky='e', padx=5, pady=5)
    
    
    
    def export_selected_files():
        selected_files = [filename for filename, var in main_files_checkboxes.items() if var.get() == 1] + [filename for filename, var in supplementary_files_checkboxes.items() if var.get() == 1]
        extract_files_nodelete(selected_files, lib_directory,output_directory)
        reset_action_flag()
    
    export_selected_button = tk.Button(frame3, text="Selected", command=export_selected_files)
    export_selected_button.grid(row=1, column=2, sticky='e', padx=5, pady=5)

    # export_selected_button = tk.Button(frame3, text="Selected", command=lambda: callback(open_selected_files()))

    def export_all_files():
        all_files = all_m_files + all_s_files
        extract_files_nodelete(all_files, lib_directory,output_directory)
        reset_action_flag()

    export_all_button=tk.Button(frame3, text="All", command=export_all_files)
    export_all_button.grid(row=1, column=3, sticky='e', padx=5, pady=5)
    
    # open_all_button = tk.Button(frame3, text="All", command=lambda: callback(open_all_files()))


    exit_button = tk.Button(frame3, text="Exit (esc)", command=reset_action_flag)
    exit_button.grid(row=2, column=3, sticky='e', padx=5, pady=5)
    new_window.bind("<Escape>", lambda _: exit_button.invoke())


    frame3.columnconfigure(0, weight=1)

    center_window(new_window)
    new_window.deiconify() # Show the window
    # Bind the window close event to reset_action_flag
    new_window.protocol("WM_DELETE_WINDOW", reset_action_flag)


def minimize_window():
    root.iconify()  # Minimize the window

def on_closing():
    minimize_window()  # Minimize on "X" button click

def exit_program():
    root.quit()
    root.destroy()  # Close the Tkinter window
    exit_flag.set()   # Exit the whole code

def a_button_openadv():
    global action_in_progress
    if not action_in_progress:
        # Start a new thread to run the specific code
        thread = threading.Thread(target=a_openadv)
        thread.start()

# Create a Toplevel window

# Position the window in the bottom-right corner
window_width = 500
window_height = 300
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
yadjust = 100  # Adjust this value as needed
xadjust = 10
position_top = screen_height - window_height - yadjust
position_right = screen_width - window_width - xadjust
root.geometry(f"{window_width}x{window_height}+{position_right}+{position_top}")

# Create three frames
frame1 = tk.Frame(root, relief="solid", padx=5, pady=5)
frame2 = tk.Frame(root, relief="solid", padx=5, pady=5)
frame3 = tk.Frame(root, relief="solid", padx=5, pady=5)

frame1.pack(side="top", fill="both", expand=True)
frame2.pack(side="top", fill="both", expand=True)
frame3.pack(side="top", fill="both", expand=True)

# Add three buttons to frame1
button_setting = tk.Button(frame1, text="")
button_outputfolder = tk.Button(frame1, text="Open output folder",command=force_open_output)
button_libfolder = tk.Button(frame1, text="Open library folder",command=force_open_lib)

button_setting.pack(side="right", padx=5)
button_outputfolder.pack(side="right", padx=5)
button_libfolder.pack(side="right", padx=5)

# Add labels and buttons to frame2
label1_1 = tk.Label(frame2, text="From Excel selection:", anchor="w")
label1_1.grid(row=0, column=0, padx=5, pady=5, sticky="w")

from excel_check_file_exist import check_file_exist
from excel_input_pubmed_data import input_pubmed_data

def input_pubmed_data2():
    global action_in_progress
    action_in_progress = True  # Set the action in progress
    check_file_exist(lib_directory)
    input_pubmed_data()
    action_in_progress = False
button_1 = tk.Button(frame2, text="Import PubMed Data",command=input_pubmed_data2)
button_1.grid(row=0, column=1, padx=5, pady=5, sticky="w")

def check_file_exist2():
    global action_in_progress
    action_in_progress = True  # Set the action in progress
    check_file_exist(lib_directory)
    action_in_progress = False
button_2 = tk.Button(frame2, text="Check if files exist",command=check_file_exist2)
button_2.grid(row=1, column=1, padx=5, pady=5, sticky="w")

# button_3 = tk.Button(frame2, text="Inspect files")
# button_3.grid(row=2, column=1, padx=5, pady=5, sticky="w")

label1_5 = tk.Label(frame2, text="From clipboard:", anchor="w")
label1_5.grid(row=4, column=0, padx=5, pady=5, sticky="w")

button_5 = tk.Button(frame2, text="Open files", command = a_openfile)
button_5.grid(row=4, column=1, padx=5, pady=5, sticky="w")
label2_5 = tk.Label(frame2, text="Ctrl + Shift + K", anchor="e")
label2_5.grid(row=4, column=3, padx=5, pady=5, sticky="e")

button_6 = tk.Button(frame2, text="Inspect files", command=a_button_openadv)
button_6.grid(row=5, column=1, padx=5, pady=5, sticky="w")
label2_6 = tk.Label(frame2, text="Ctrl + Shift + J", anchor="e")
label2_6.grid(row=5, column=3, padx=5, pady=5, sticky="e")

# Make the third column expandable
frame2.columnconfigure(2, weight=1)

# Repeat the above for rows 2 to 6

# Add a button to frame3
exit_button = tk.Button(frame3, text="Exit", command=exit_program)
exit_button.pack(side="right", padx=5)
    
# Set the window close event handler to minimize on "X" button click
root.protocol("WM_DELETE_WINDOW", on_closing)

# Create an exit flag to signal the main loop to terminate
exit_flag = threading.Event()

def check_shortcut():
    global action_in_progress

    while not exit_flag.is_set():
        # Listen for the Ctrl+Shift+J shortcut
        if keyboard.is_pressed('ctrl+shift+j'):
            if not action_in_progress:
                # Start a new thread to run the specific code
                thread = threading.Thread(target=a_openadv)
                thread.start()

            # Listen for the Ctrl+Shift+J shortcut
        if keyboard.is_pressed('ctrl+shift+k'):
            if not action_in_progress:
                # Start a new thread to run the specific code
                thread = threading.Thread(target=a_openfile)
                thread.start()
        
        time.sleep(0.05)


# Start the shortcut detection in the main thread
shortcut_thread = threading.Thread(target=check_shortcut)
shortcut_thread.start()

if __name__ == "__main__":
    root.mainloop()
