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
    This is the main() function from the original script
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
        show_feedback("Merge completed successfully.", "green")

# Merges two specific PDF files
def manual_merge_pdfs(file1, file2, output_filename, output_folder):
    
    # Create merger object/list
    merger = PdfWriter()
    for file in [file1, file2]:
        # Add file to merger object
        merger.append(file)

    # Form output path of new merged file
    output_path = os.path.join(output_folder, output_filename)

    # Write new file to output folder and close
    merger.write(output_path)
    merger.close()
    show_feedback("Merge completed successfully.", "green")
    
# Function for input folder "Browse..." button
def browse_input_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        auto_folder_path_var.set(folder_selected)
        folder_field.xview_moveto(1)    # Show end of path

# Function for output folder "Browse..." button
def browse_output_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        output_path_var.set(folder_selected)
        folder_field.xview_moveto(1)    # Show end of path

# Function for choosing input PDF files
def browse_file(pdf_number):
    file_selected = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]  # Criteria for valid file
    )

    if file_selected:
        if pdf_number == "pdf1":
            pdf1_path_var.set(file_selected)
            chosen_pdf_1.xview_moveto(1)    # Show end of path
        elif pdf_number == "pdf2":
            pdf2_path_var.set(file_selected)
            chosen_pdf_2.xview_moveto(1)    # Show end of path

# Function for "Merge PDFs" button (calls main functions)
def on_merge_click():
    if validate_inputs():
        if radio_selection.get() == "auto_merge":
            input_folder = auto_folder_path_var.get()
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

# Checks whether file is PDF
def is_pdf_file(filepath):
    return os.path.isfile(filepath) and filepath.lower().endswith(".pdf")

# Function called whenever radio button selection changes
def on_radio_option_change():
    # Clear status message
    status_var.set("")

    # Clear field values
    for field in [auto_folder_path_var, output_path_var, pdf1_path_var, pdf2_path_var]:
        field.set("")
    
    if radio_selection.get() == "auto_merge":
        # Enable auto functionality
        folder_field.config(state="normal")
        browse_button.config(state="normal")
        output_browse_button.config(state="normal")

        # Disable manual functionality
        chosen_pdf_1.config(state="disabled")
        chosen_pdf_2.config(state="disabled")
        output_name.config(state="disabled")
        browse_pdf1_button.config(state="disabled")
        browse_pdf2_button.config(state="disabled")
    else:
        # Enable manual functionality
        chosen_pdf_1.config(state="normal")
        chosen_pdf_2.config(state="normal")
        output_name.config(state="normal")
        output_folder_field.config(state="normal")
        browse_pdf1_button.config(state="normal")
        browse_pdf2_button.config(state="normal")
        output_browse_button.config(state="normal")

        # Disable auto functionality
        browse_button.config(state="disabled")
        folder_field.config(state="disabled")

# Checks for inputs to enable Merge
def check_for_inputs(*args):    # Must accept automatically passed arguments
    # ______ Check Auto Inputs ______
    if radio_selection.get() == "auto_merge":
        if auto_folder_path_var.get() and output_path_var.get():
            merge_button.config(state="normal")  # Enable button
            show_feedback("Ready", "green")
        else:
            merge_button.config(state="disabled")  # Disable button
    
    # ______ Check Manual Inputs ______
    else:
        # Validate PDF file inputs
        if pdf1_path_var.get() and pdf2_path_var.get() and output_filename_var.get() and output_path_var.get():
            merge_button.config(state="normal")  # Enable button
            show_feedback("Ready", "green")
        else:
            merge_button.config(state="disabled")  # Disable button

# Validates inputs
def validate_inputs():
    '''
    Called within multiple functions whenever any user inputs change.
    Existence of an input is verified before calling validate_inputs().
    '''
    valid_inputs = False

    # ______ Auto Inputs ______
    # Validate folder with PDF files
    if radio_selection.get() == "auto_merge":
        if os.path.isdir(auto_folder_path_var.get()) and folder_contains_pdfs(auto_folder_path_var.get()):
            valid_inputs = True
        else:
            show_feedback("Folder does not contain PDF files", "red")
            valid_inputs = False
    # ______ Manual Inputs ______
    else:
        valid_pdfs = False
        valid_output_filename = False
        valid_output_folder = False

        # Validate PDF file inputs
        if is_pdf_file(pdf1_path_var.get()) and is_pdf_file(pdf2_path_var.get()):
            valid_pdfs = True
        elif not is_pdf_file(pdf1_path_var.get()) and is_pdf_file(pdf2_path_var.get()):
            valid_pdfs = False
            show_feedback("Invalid file selection - verify that PDF #1 is a PDF file", "red")
        elif is_pdf_file(pdf1_path_var.get()) and not is_pdf_file(pdf2_path_var.get()):
            valid_pdfs = False
            show_feedback("Invalid file selection - verify that PDF #2 is a PDF file", "red")
        else:
            valid_pdfs = False
            show_feedback("Invalid file selections - verify that chosen files are PDF files", "red")

        # Validate output filename
        if output_filename_var.get()[-4:] == ".pdf":
            valid_output_filename = True
        else:
            valid_output_filename = False
            show_feedback("Output filename must end with .pdf", "red")
        
        # Combine individual validations
        valid_inputs = valid_pdfs and valid_output_filename
    
    # ______ Auto & Manual Inputs ______
    # Validate folder for output file
    if os.path.isdir(output_path_var.get()):
        valid_output_folder = True
    else:
        valid_output_folder = False
        show_feedback("Invalid output folder selection", "red")

    return valid_inputs and valid_output_folder


# Feedback helper function for varying message and color
def show_feedback(message, color):
    status_var.set(message) # Updates message
    status_label.config(fg=color)   # Updates text color



# __________ GUI Setup __________ #
root = tk.Tk()
root.title("PDF Merger")
root.geometry("800x450")    # window width & height in pixels

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
auto_folder_path_var = tk.StringVar()    
auto_folder_path_var.trace_add("write", check_for_inputs)    # validate folder selection whenever it changes

# Holds radio button selection
radio_selection = tk.StringVar(value="none")  # Default to "none" to have invisible radio selected (visual glitch otherwise)

# Holds PDF #1 selection
pdf1_path_var = tk.StringVar()    
pdf1_path_var.trace_add("write", check_for_inputs)    # validate folder selection whenever it changes

# Holds PDF #2 selection
pdf2_path_var = tk.StringVar()    
pdf2_path_var.trace_add("write", check_for_inputs)    # validate folder selection whenever it changes

# Holds output filename
output_filename_var = tk.StringVar()
output_filename_var.trace_add("write", check_for_inputs)

# Holds output folder path string
output_path_var = tk.StringVar()
output_path_var.trace_add("write", check_for_inputs)    # validate folder selection whenever it changes

# Holds status message string
status_var = tk.StringVar() 


# __________ HEADER GUI ELEMENTS __________ #

# Tool overview/introduction
intro_text = (
    "Use this tool to merge PDF files. It is intended for two different uses:"
    "\n    1) automatically merge resolution and stipulation PDFs for the SLA process"
    "\n    2) manually choose and merge two specific PDF files"
)
introduction = tk.Label(root, text=intro_text, font=("Helvetica", 10), wraplength=700, justify="left")
introduction.grid(row=0, column=0, columnspan=2, sticky="w", pady=10, padx=10)

# __________ LEFT FRAME GUI ELEMENTS __________ #

# Radio button section
# AUTO
instructions = tk.Label(left_frame, text="Choose an option:", font=("Helvetica", 12, "bold"))
instructions.grid(row=1, column=0, sticky="w", pady=(0, 5))

auto_radio_button = tk.Radiobutton(left_frame, text="Auto SLA PDF Merge", variable=radio_selection, value="auto_merge", command=on_radio_option_change, font=("Helvetica", 10))
auto_radio_button.grid(row=2, column=0, sticky="w", pady=5)
auto_explainer_text = (
    "NOTE: requires input folder with prepared PDF files (see SLA process guide) and output folder"
)
auto_instruction = tk.Label(left_frame, text=auto_explainer_text, font=("Helvetica", 8), wraplength=150, justify="left")
auto_instruction.grid(row=3, column=0, sticky="w", pady=5, padx=(5,0))

# MANUAL
manual_radio_button = tk.Radiobutton(left_frame, text="Manual PDF Merge", variable=radio_selection, value="manual_merge", command=on_radio_option_change, font=("Helvetica", 10))
manual_radio_button.grid(row=4, column=0, sticky="w", pady=5)
manual_explainer_text = (
    "NOTE: requires two input files, output filename, and output folder"
)
manual_instruction = tk.Label(left_frame, text=manual_explainer_text, font=("Helvetica", 8), wraplength=150, justify="left")
manual_instruction.grid(row=5, column=0, sticky="w", pady=5, padx=(5,0))

# Invisible radio button required to create initial view with no selections
dummy_radio_button = tk.Radiobutton(left_frame, variable=radio_selection, value="none")
dummy_radio_button.pack_forget()    # Makes button invisible

# Label for general status feedback (ex: Success, Error, etc.)
status_label = tk.Label(left_frame, textvariable=status_var, font=("Helvetica", 12, "bold"), wraplength=150, justify="center")
status_label.grid(row=6, column=0, pady=(30,0))


# __________ RIGHT FRAME GUI ELEMENTS __________ #

# _____ AUTO MERGE FUNCTIONALITY _____ #
# Auto merge instructions
auto_label = tk.Label(right_frame, text="Select folder with PDFs:")
auto_label.grid(row=1, column=0, padx=(10,10), pady=(40,75))

# Create field showing chosen folder
folder_field = tk.Entry(right_frame, textvariable=auto_folder_path_var, width=50, state="disabled")
folder_field.grid(row=1, column=1, padx=(10,10), pady=(40,75))

# Create browse button
browse_button = tk.Button(right_frame, text="Browse...", command=browse_input_folder, state="disabled")
browse_button.grid(row=1, column=2, padx=(10,0), pady=(40,75))


# _____ MANUAL MERGE FUNCTIONALITY _____ #
manual_base_row = 4

pdf1_label = tk.Label(right_frame, text="Select PDF #1:")
pdf1_label.grid(row=manual_base_row, column=0, padx=(10,10), pady=(5,5))

chosen_pdf_1 = tk.Entry(right_frame, textvariable=pdf1_path_var, width=50, state="disabled")
chosen_pdf_1.grid(row=manual_base_row, column=1)

browse_pdf1_button = tk.Button(right_frame, text="Choose PDF #1...", command=lambda: browse_file("pdf1"), state="disabled") # lambda prevents immediate run of function with parameters
browse_pdf1_button.grid(row=manual_base_row, column=2)

pdf2_label = tk.Label(right_frame, text="Select PDF #2:")
pdf2_label.grid(row=manual_base_row+1, column=0, padx=(10,10), pady=(5,5))

chosen_pdf_2 = tk.Entry(right_frame, textvariable=pdf2_path_var, width=50,  state="disabled")
chosen_pdf_2.grid(row=manual_base_row+1, column=1)

browse_pdf2_button = tk.Button(right_frame, text="Choose PDF #2...", command=lambda: browse_file("pdf2"), state="disabled") # lambda prevents immediate run of function with parameters
browse_pdf2_button.grid(row=manual_base_row+1, column=2)

output_name_label = tk.Label(right_frame, text="Enter output filename:")
output_name_label.grid(row=manual_base_row+2, column=0, padx=(10,10), pady=(5,5))

output_name = tk.Entry(right_frame, textvariable=output_filename_var, width=50, state="disabled")
output_name.grid(row=manual_base_row+2, column=1)


# _____ SHARED FUNCTIONALITY _____ #

# Output folder instructions
output_folder_label = tk.Label(right_frame, text="Select output folder:")
output_folder_label.grid(row=manual_base_row+4, column=0, pady=(50,10))

# Create field showing chosen folder
output_folder_field = tk.Entry(right_frame, textvariable=output_path_var, width=50, state="disabled")
output_folder_field.grid(row=manual_base_row+4, column=1, pady=(50,10))

# Create output browse button
output_browse_button = tk.Button(right_frame, text="Browse...", command=browse_output_folder, state="disabled")
output_browse_button.grid(row=manual_base_row+4, column=2, pady=(50,10))

# Button to run merge
merge_button = tk.Button(right_frame, text="Merge PDFs", command=on_merge_click, state="disabled")
merge_button.grid(row=manual_base_row+5, column=1, pady=(10,10))

# ____________________________________________________

# Runs GUI window
root.mainloop()
