'''
Description:
    This program creates a user interface to merge PDF files, specifically Community Board resolutions
    and SLA stipulation documents. It can also be used for general PDF merges.

Auto SLA Workflow:
1) Create PDFs of resolutions and save in "Resolution Document Components" Sharepoint folder
2) Save all stipulation PDFs in "Resolution Document Components" folder
3) Delete contents of working folders: "Resolution Document Components" and "SLA Resolutions wSTIPS"
3) Download the entire "Resolution Document Components" folder (it will be zipped)
4) Extract zipped/compressed folder contents to a new folder
5) Open PDF Merger interface and use automatic merge function
6) Check output files for accuracy
7) Upload all output files to correct Sharepoint folder

'''

import os
import tkinter as tk
from tkinter import filedialog, messagebox
import re
from pypdf import PdfWriter

# Merges all pairable PDF files based on numbering criteria
def auto_merge_pdfs(selected_folder, output_folder):
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
            if filename.endswith('.pdf') and "cb3 reso" in filename.lower() and re.match(r"\s*\d+", filename):
                resolutions.append(filename)
            # Stipulation documents lack "cb3 reso" text
            elif filename.endswith('.pdf') and "cb3 reso" not in filename.lower() and re.match(r"\s*\d+", filename):
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

    # Input and output folders come from GUI
    input_folder = selected_folder
    output_folder = output_folder

    # Make pairs of the files to merge them
    pairs = pair_files(input_folder)

    # Verify that functions found valid pairs to merge
    if pairs:
        print(pairs)
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

# Merges two specific PDF files
def manual_merge_pdfs(file1, file2, output_filename, output_folder):
    
    # Create merger object/list
    merger = PdfWriter()
    for file in [file1, file2]:
        # # Form file path from existing file name
        # file_path = os.path.join(output_folder, file)
        # Add file to merger object
        merger.append(file)

    # Edit filename
    # output_filename = p[0].replace(".pdf", " wSTIPS.pdf")

    # Form output path of new merged file
    output_path = os.path.join(output_folder, output_filename)

    # Write new file to output folder and close
    merger.write(output_path)
    merger.close()
    
# Function for input folder "Browse..." button
def browse_input_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_path_var.set(folder_selected)

        # Update size of field box based on file path size
        # folder_field.config(width=max(50, len(folder_selected)+10))
        folder_field.xview_moveto(1)    # Show end of path

# Function for output folder "Browse..." button
def browse_output_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        output_path_var.set(folder_selected)

        # Update size of field box based on file path size
        # folder_field.config(width=max(50, len(folder_selected)+10))
        folder_field.xview_moveto(1)    # Show end of path

# Function for choosing input PDF files
def browse_file(pdf_number):
    file_selected = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )

    if file_selected:
        if pdf_number == "pdf1":
            pdf1_path_var.set(file_selected)
            # chosen_pdf_1.config(width=max(50, len(file_selected)+10))
            chosen_pdf_1.xview_moveto(1)    # Show end of path
        elif pdf_number == "pdf2":
            pdf2_path_var.set(file_selected)
            # chosen_pdf_2.config(width=max(50, len(file_selected)+10))
            chosen_pdf_2.xview_moveto(1)    # Show end of path

# Function for "Merge PDFs" button (calls main function)
def on_merge_click():
    if radio_selection.get() == "auto_merge":
        input_folder = folder_path_var.get()
        output_folder = output_path_var.get()
        auto_merge_pdfs(input_folder, output_folder)
    else:
        file1 = pdf1_path_var.get()
        file2 = pdf2_path_var.get()
        output_filename = output_filename_var.get()
        output_folder = output_path_var.get()
        manual_merge_pdfs(file1, file2, output_filename, output_folder)
    
# Function to check if folder contains any PDFs
def folder_contains_pdfs(folder):
    try:
        # List files and check if any end with .pdf (case insensitive)
        files = os.listdir(folder)
        return any(f.lower().endswith(".pdf") for f in files)
    except Exception:
        # If folder is not accessible or any error occurs, treat as invalid
        return False

# Function called whenever folder path changes
def on_folder_change(*args):
    ready_state = False
    auto_input_folder = folder_path_var.get()
    output_folder = output_path_var.get()
    print(output_folder)
    # Enable the Merge button only if:
    # 1. The folder exists and is a directory
    # 2. The folder contains at least one PDF file
    if radio_selection.get() == "auto_merge":
        if os.path.isdir(auto_input_folder) and folder_contains_pdfs(auto_input_folder):
            merge_button.config(state="normal")  # Enable button
            status_var.set("Ready")
        else:
            merge_button.config(state="disabled")  # Disable button
            status_var.set("Invalid folder selection - verify choice is a valid folder with PDF files.")
    else:
        if os.path.isdir(output_folder):
            merge_button.config(state="normal")  # Enable button
            status_var.set("Ready")
        else:
            merge_button.config(state="disabled")  # Disable button
            status_var.set("Invalid output folder selection.")

# Checks whether file is PDF
def is_pdf_file(filepath):
    return os.path.isfile(filepath) and filepath.lower().endswith(".pdf")

# Validates file selection whenever manual merge file is selected
def on_file_change(*args):
    file1 = pdf1_path_var.get()
    file2 = pdf2_path_var.get()

    if is_pdf_file(file1) and is_pdf_file(file2):
        merge_button.config(state="normal")  # Enable button
        status_var.set("Ready")
    elif not is_pdf_file(file1) and is_pdf_file(file2):
        merge_button.config(state="disabled")  # Disable button
        status_var.set("Invalid file selection - verify that PDF #1 is a PDF file.")
    elif is_pdf_file(file1) and not is_pdf_file(file2):
        merge_button.config(state="disabled")  # Disable button
        status_var.set("Invalid file selection - verify that PDF #2 is a PDF file.")
    else:
        merge_button.config(state="disabled")  # Disable button
        status_var.set("Invalid file selection - verify that chosen files are PDF files.")

# Function called whenever radio button selection changes
def on_radio_option_change():
    # Clear status message
    status_var.set("")

    # Clear field values
    for field in [folder_path_var, output_path_var, pdf1_path_var, pdf2_path_var]:
        field.set("")
    
    if radio_selection.get() == "auto_merge":
        print("auto merge")

        # Enable auto functionality
        browse_button.config(state="normal")

        # Disable manual functionality
        browse_pdf1_button.config(state="disabled")
        browse_pdf2_button.config(state="disabled")
    else:
        print("manual merge")

        # Enable manual functionality
        browse_pdf1_button.config(state="normal")
        browse_pdf2_button.config(state="normal")

        # Disable auto functionality
        browse_button.config(state="disabled")

# GUI Setup
root = tk.Tk()
root.title("PDF Merger")
root.geometry("750x500")    # window width & height in pixels

# Configure column weights so column 1 (right side) expands
root.columnconfigure(0, weight=0)  # Left column (radio buttons) - fixed
root.columnconfigure(1, weight=1)  # Right column (main content) - flexible

# Header section goes at row=0

# Left frame for radio buttons
left_frame = tk.Frame(root)
left_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)

# Right frame for other functionality
right_frame = tk.Frame(root, bg="lightgray")
right_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)


# __________ GUI VARIABLES __________ #

# Holds auto merge folder path string
folder_path_var = tk.StringVar()    
folder_path_var.trace_add("write", on_folder_change)    # validate folder selection whenever it changes

# Holds status message string
status_var = tk.StringVar() 

# Holds radio button selection
radio_selection = tk.StringVar(value="none")  # Default to "none" to have invisible radio selected (visual glitch otherwise)

# Holds PDF #1 selection
pdf1_path_var = tk.StringVar()    
pdf1_path_var.trace_add("write", on_file_change)    # validate folder selection whenever it changes

# Holds PDF #2 selection
pdf2_path_var = tk.StringVar()    
pdf2_path_var.trace_add("write", on_file_change)    # validate folder selection whenever it changes

output_filename_var = tk.StringVar()

# Holds output folder path string
output_path_var = tk.StringVar()
output_path_var.trace_add("write", on_folder_change)    # validate folder selection whenever it changes

# __________________________________ #

'''
HEADER SECTION
'''
# Tool overview/introduction
intro_text = (
    "Use this tool to merge PDF files. It is intended for two different uses:"
    "\n    1) automatically merge resolution and stipulation PDFs for the SLA process"
    "\n    2) merge two specific PDF files"
)
introduction = tk.Label(root, text=intro_text, font=("Helvetica", 10), wraplength=700, justify="left")
introduction.grid(row=0, column=0, columnspan=2, sticky="w", pady=10, padx=10)

'''
LEFT FRAME SECTION
'''

# Radio button section
# AUTO
instructions = tk.Label(left_frame, text="Choose an option:", font=("Helvetica", 12, "bold"))
instructions.grid(row=1, column=0, sticky="w", pady=(10, 5))

auto_radio_button = tk.Radiobutton(left_frame, text="SLA Auto PDF Merge", variable=radio_selection, value="auto_merge", command=on_radio_option_change, font=("Helvetica", 10))
auto_radio_button.grid(row=2, column=0, sticky="w", pady=5)
auto_explainer_text = (
    "NOTE: requires input folder with prepared PDF files (see SLA process guide) and output folder"
)
auto_instruction = tk.Label(left_frame, text=auto_explainer_text, font=("Helvetica", 8), wraplength=150, justify="left")
auto_instruction.grid(row=3, column=0, sticky="w", pady=5)

# MANUAL
manual_radio_button = tk.Radiobutton(left_frame, text="Manual PDF Merge", variable=radio_selection, value="manual_merge", command=on_radio_option_change, font=("Helvetica", 10))
manual_radio_button.grid(row=4, column=0, sticky="w", pady=5)
manual_explainer_text = (
    "NOTE: requires two input files, output filename, and output folder"
)
manual_instruction = tk.Label(left_frame, text=manual_explainer_text, font=("Helvetica", 8), wraplength=150, justify="left")
manual_instruction.grid(row=5, column=0, sticky="w", pady=5)

# Invisible radio button required to create initial view with no selections
dummy_radio_button = tk.Radiobutton(left_frame, variable=radio_selection, value="none")
dummy_radio_button.pack_forget()    # Makes button invisible

# Label for general status feedback (ex: Success, Error, etc.)
status_label = tk.Label(left_frame, textvariable=status_var, font=("Helvetica", 12, "bold"), fg="red", wraplength=150, justify="center")
status_label.grid(row=6, column=0)


'''
RIGHT FRAME SECTION
'''
# _____ AUTO MERGE _____ #
# Auto merge instructions
auto_label = tk.Label(right_frame, text="Select folder with PDFs:")
auto_label.grid(row=1, column=0)

# Create field showing chosen folder
folder_field = tk.Entry(right_frame, textvariable=folder_path_var, width=50)
folder_field.grid(row=1, column=1)

# Create browse button
browse_button = tk.Button(right_frame, text="Browse...", command=browse_input_folder, state="disabled")
browse_button.grid(row=1, column=2)


# _____ MANUAL MERGE _____ #
manual_base_row = 4

# Create browse button & display to show chosen file
browse_pdf1_button = tk.Button(right_frame, text="Choose PDF #1...", command=lambda: browse_file("pdf1"), state="disabled") # lambda prevents immediate run of function with parameters
browse_pdf1_button.grid(row=manual_base_row, column=0)

chosen_pdf_1 = tk.Entry(right_frame, textvariable=pdf1_path_var, width=50)
chosen_pdf_1.grid(row=manual_base_row, column=1)

browse_pdf2_button = tk.Button(right_frame, text="Choose PDF #2...", command=lambda: browse_file("pdf2"), state="disabled") # lambda prevents immediate run of function with parameters
browse_pdf2_button.grid(row=manual_base_row+1, column=0)

chosen_pdf_2 = tk.Entry(right_frame, textvariable=pdf2_path_var, width=50)
chosen_pdf_2.grid(row=manual_base_row+1, column=1)

output_name_label = tk.Label(right_frame, text="Enter an output filename:")
output_name_label.grid(row=manual_base_row+2, column=0)

output_name = tk.Entry(right_frame, textvariable=output_filename_var, width=50)
output_name.grid(row=manual_base_row+2, column=1)

# _____ SHARED FUNCTIONALITY _____ #

# Output folder instructions
output_folder_label = tk.Label(right_frame, text="Select an output folder:")
output_folder_label.grid(row=manual_base_row+4, column=0)

# Create field showing chosen folder
output_field = tk.Entry(right_frame, textvariable=output_path_var, width=50)
output_field.grid(row=manual_base_row+4, column=1)

# Create output browse button
output_browse_button = tk.Button(right_frame, text="Browse...", command=browse_output_folder, state="normal")
output_browse_button.grid(row=manual_base_row+4, column=2)

# Button to run merge
merge_button = tk.Button(right_frame, text="Merge PDFs", command=on_merge_click, state="disabled")
merge_button.grid(row=manual_base_row+5, column=1)

# ____________________________________________________
# Runs GUI window
root.mainloop()
