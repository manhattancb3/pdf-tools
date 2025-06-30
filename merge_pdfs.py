'''
Description:
    This script is intended to merge PDF files, specifically Community Board resolutions
    and SLA stipulation documents. It can also be used for Cannabis resolution and stipulations.




Potential updates:
- create GUI for choosing input folder
- create separate merge function
'''

# Import statements - make sure you pip install if running on new machine
import os
import re
from pypdf import PdfWriter

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
    
    pairs = []
    for r in resolutions:
        for s in stipulations:
            if extract_number(r) == extract_number(s):
                pairs.append([r,s]) # Append pair list to main list
    
    return pairs


def main():
    '''

    '''
    # Input folder MUST exist, output folder created automatically
    input_folder = r"C:\Users\MN03\Documents\Python Scripts\PDF Tools\Resolution Document Components"
    output_folder = r"C:\Users\MN03\Documents\Python Scripts\PDF Tools\SLA Resolutions wSTIPS"
    os.makedirs(output_folder, exist_ok=True)   # Create folder if it doesn't exist

    # Make pairs of the files to merge them
    pairs = pair_files(input_folder)

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

main()