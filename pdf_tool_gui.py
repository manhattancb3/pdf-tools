'''
Description:
    This script is intended to merge PDF files, specifically Community Board resolutions
    and SLA stipulation documents. It can also be used for Cannabis resolution and stipulations.


Potential updates:
- create separate merge function for general use (not specific to SLA process)

Workflow:
1) Create PDFs of resolutions and save in "Resolution Document Components" Sharepoint folder
2) Save all stipulation PDFs in "Resolution Document Components" folder
3) Delete contents of working folders: "Resolution Document Components" and "SLA Resolutions wSTIPS"
3) Download the entire "Resolution Document Components" folder (it will be zipped)
4) Extract zipped/compressed folder contents to "PDF Tools" local folder (where this script is stored)
5) Run script
6) Check output files for accuracy
7) Upload all output files to correct Sharepoint folder


GUI Design
- choose manual operation
    - pick PDFs to combine and order
- choose automatic operation
    - pick specific folder and verify matching criteria
'''

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from pypdf import PdfWriter

def auto_merge_pdfs(selected_folder):
    '''
    This is the main() function from the original 
    '''
    # Finds leading numbers in filenames
    def extract_number(filename):
        '''
        Uses regular expressions to extract the number at the beginning of filenames.
        The extracted number is used to match pairs of PDF documents.

        Returns an integer or None.
        '''
        match = re.match(r"\s*\d+", filename)
        if match:
            return int(match.group())
        else:
            print("\n@@@ --- WARNING: filename missing the leading number --- @@@\n")
            return None
    
    # Pairs files with same number label
    def pair_files(directory):
        '''
        Loops over Windows folder and finds all resolution and stipulation documents.
        Creates matches using the extracted number and returns a list of pairs (nested list).

        Returns a list.
        '''
        resolutions = []
        stipulations = []
        for filename in os.listdir(directory):
            # Resolution documents should always contain "cb3 reso" text
            if filename.endswith('.pdf') and "cb3 reso" in filename.lower():
                resolutions.append(filename)
            # Stipulation documents lack "cb3 reso" text
            elif filename.endswith('.pdf') and "cb3 reso" not in filename.lower():
                stipulations.append(filename)
        
        # Check whether resolution AND stipulation files were found
        if not(resolutions and stipulations):
            # Return None to trigger error message
            return None
        else:
            pairs = []
            for r in resolutions:
                for s in stipulations:
                    if extract_number(r) == extract_number(s):
                        pairs.append([r,s]) # Append pair list to main list
            return pairs

    # Input folder MUST exist, output folder created automatically
    input_folder = selected_folder
    output_folder = r"C:\Users\MN03\Documents\Python Scripts\PDF Tools\SLA Resolutions wSTIPS"
    os.makedirs(output_folder, exist_ok=True)   # Create folder if it doesn't exist

    # Make pairs of the files to merge them
    pairs = pair_files(input_folder)

    # Verify that functions found valid pairs to merge
    if pairs:
        for p in pairs:
            # Create merger object/list
            merger = PdfWriter()
            for file in p:
                # Form file path from existing file name
                file_path = os.path.join(input_folder, file)
                # Add file to merger object
                merger.append(file_path)

            # Edit filename
            output_filename = p[0].replace(".pdf", " wSTIPS.pdf")

            # Form output path of new merged file
            output_path = os.path.join(output_folder, output_filename)

            # Write new file to output folder and close
            merger.write(output_path)
            merger.close()
        
        # Send feedback message once merge is completed
        status_var.set("Merge completed successfully.")
    else:
        # Send error message
        status_var.set("Error: folder does not contain files valid for auto-merge")

    
# Function for "Browse..." button
def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_path_var.set(folder_selected)

        # Update size of field box based on file path size
        folder_field.config(width=max(50, len(folder_selected)+10))

# Function for "Merge PDFs" button (calls main function)
def on_merge_click():

    folder = folder_path_var.get()
    auto_merge_pdfs(folder)

# Function to check if folder contains any PDFs
def folder_contains_pdfs(folder):
    try:
        # List files and check if any end with .pdf (case insensitive)
        files = os.listdir(folder)
        return any(f.lower().endswith(".pdf") for f in files)
    except Exception:
        # If folder is not accessible or any error occurs, treat as invalid
        return False

# This function is called every time the folder path changes
def on_folder_change(*args):
    folder = folder_path_var.get()
    # Enable the Merge button only if:
    # 1. The folder exists and is a directory
    # 2. The folder contains at least one PDF file
    if os.path.isdir(folder) and folder_contains_pdfs(folder):
        merge_button.config(state="normal")  # Enable button
        status_var.set("Ready")
    else:
        merge_button.config(state="disabled")  # Disable button
        status_var.set("Invalid folder selection - verify that folder contains PDFs.")


def on_radio_option_change():
    if radio_selection.get() == "auto_merge":
        print("auto merge")
    else:
        print("manual merge")

# GUI Setup
root = tk.Tk()
root.title("PDF Merger")
root.geometry("750x500")    # window width & height in pixels

# # Configure column weights so column 1 (right side) expands
# root.columnconfigure(0, weight=0)  # Left column (radio buttons) - fixed
# root.columnconfigure(1, weight=1)  # Right column (main content) - flexible

# # Left frame for radio buttons
# left_frame = tk.Frame(root)
# left_frame.grid(row=0, column=0, sticky="ns", padx=10, pady=10)

# # Right frame for other functionality
# right_frame = tk.Frame(root, bg="lightgray")
# right_frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)




# Containers for GUI variables
folder_path_var = tk.StringVar()    # Holds folder path string
folder_path_var.trace_add("write", on_folder_change)    # validate folder selection whenever it changes

status_var = tk.StringVar() # Holds status message string

radio_selection = tk.StringVar(value="auto_merge")  # Holds radio button selection


# Radio button panel
# tk.Label(root, text="Choose an option:").pack()
auto_radio_button = tk.Radiobutton(root, text="SLA Auto Merge PDFs", variable=radio_selection, value="auto_merge", command=on_radio_option_change)
auto_radio_button.pack(side="left")
manual_radio_button = tk.Radiobutton(root, text="Manual Merge PDFs", variable=radio_selection, value="manual_merge", command=on_radio_option_change)
manual_radio_button.pack()

tk.Label(root, text="Select folder with PDFs:").pack(padx=10, pady=5)

# Create field showing chosen folder
folder_field = tk.Entry(root, textvariable=folder_path_var, width=50)
folder_field.pack(padx=10, pady=5)  

# Create browse button
tk.Button(root, text="Browse...", command=browse_folder).pack(padx=10, pady=5)

# Button to run automatic merge
merge_button = tk.Button(root, text="Merge PDFs", command=on_merge_click, state="disabled")
merge_button.pack(padx=10, pady=10)


# Label for general status feedback (ex: Success, Error, etc.)
tk.Label(root, textvariable=status_var).pack()
root.mainloop()