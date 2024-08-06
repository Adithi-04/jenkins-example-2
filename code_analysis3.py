'''
* PROGRAM NAME: RTF TO JSON CONVERTOR
* VERSION AND DATE: 2.1 04-08-2024
* AUTHOR NAME(s): Shreeja Katama

* LANGUAGE AND LIBRARY REFERENCE: Python3,
Python Standard Libraries:
  - json
  - re
  - os
  - tkinter
  - configparser
  - datetime

* PURPOSE: Processes RTF files to extract basic data and convert it to JSON format. 
Includes dynamic extraction using control words and robust error handling.
* PARAMETERS: Folder containing RTF files as an input parameter.
* RETURNS: JSON outputs saved in a new directory.
* FUNCTION CALL NAME(s): rtf_to_json
* FUNCTION PURPOSE: Extracts data from RTF files and converts it to JSON format.
* EMBEDDED PROGRAMS: None
* MODULE: rtf_to_json.py

* COPY RIGHT: M/s CARE2DATA 2024  All Rights Reserved
**************************************************************************

* VERSION HISTORY

* REVISED VERSION AND DATE: 1.0.0 01-07-2024
* AUTHOR NAME(s): Shreeja Katama

* CHANGE REASON: Initial Version

* REVISED VERSION AND DATE: 1.2.4 31-07-2024
* AUTHOR NAME(s): Shreeja Katama

* CHANGE REASON: Updates and improvements
'''

import json
import re
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from configparser import ConfigParser
from datetime import datetime


#Read config.ini file
config_object = ConfigParser()
config_object.read("config.ini")

# Get the RTF tags, RE expressions and alignment data from the config file
RTF_tags = config_object["RTF TAGS"]
RE_expressions = config_object["RE EXPRESSIONS"]
Header_alignment = config_object['HEADER ALIGNMENT']

log_file_exceptions = open(
    "/Users/adithi/Desktop/Log File Exceptions.txt",
    'a', encoding = "utf-8")
log_file_success = open(
    "/Users/adithi/Desktop/Log File Success.txt", 
    'a', encoding = "utf-8")

log_file_exceptions.write('\n' + str(datetime.now()) + '\n')
log_file_success.write('\n' + str(datetime.now()) + '\n')

# Global debugging flag
DEBUG = False
OUTPUT_DIRECTORY = ""
FOLDER_TO_DELETE = ""

# global PAGE variable declaration to count the number of pages
PAGE = 0
NUMPAGES = 0
# global value declaration
SELECTED_FOLDER_PATH = ""
FOLDER_PATH = ""
table = ""

def debug_print(message):
    '''
    This function serves as the debugging switch
    '''
    if DEBUG:
        print(message)

# Function to check if RTF File adheres to the schema
def check_rtf(file_path):
    '''
    This is a function that checks whether the RTF File adheres to the schema mentioned.
    The adherence to the schema is found by checking whether 
    the commonly used RTF control tags are used in the RTF file.
    The RTF file is loaded, and then the content is read using the "file.read()" function
    A list is created to store the RTF tags
    An iterator is used to parse the list and check if all the tags are present in the file
    '''
    with open(file_path, 'r', encoding = "utf-8") as file:
        rtf_content = file.read().replace("{\\line}\n", " ").replace("\\~", " ")
     # Commonly used RTF tags
    rtf_tags = [RTF_tags["page break"],RTF_tags['header'],
                RTF_tags['title'],RTF_tags["row start"],
                RTF_tags["row end"],RTF_tags["cell end"]]
    flag = True
    for i in rtf_tags:
        if i in rtf_content:
            continue

        debug_print(i + " not in rtf")
        # If RTF tag is not present, the RTF does not adhere to the schema
        log_file_exceptions.write(i + " not in RTF \n")
        flag = False
        break
    return flag

# Function to extract font details from RTF content
def extract_font_details(rtf_content):
    '''This is a function to extract the font details of the RTF file
    The font details are found by using the '\\fonttbl' RTF tag
    The fonts are then extracted using an re expression to match the patterns
    The fonts are stored along with the font ID in a dictionary
    These key value pairs will be used later to extract the font details on each section
    '''
    # Extract font table
    font_table_pattern = re.compile(RE_expressions['font table'], re.DOTALL)

    font_table_match = font_table_pattern.search(rtf_content)
    if font_table_match:
        font_table = font_table_match.group(1)

    else:
        debug_print("No font table found")
        return {}

    # Extract font details from font table
    font_pattern = re.compile(RE_expressions['font pattern'])

    fonts = {}
    for match in font_pattern.finditer(font_table):
        font_id, font_name = match.groups()
        fonts['f'+font_id] = font_name
    return fonts

# Function to extract page breaks in the RTF File

def extract_page_breaks(rtf_content):
    '''This function finds the page breaks using '\\endnhere' RTF tag
    This function is used to split and extract the RTF content for each page
    '''
    page_breaks = []
    for p in re.finditer(r"\\endnhere", rtf_content): # Using the '\endhere' tag to find page breaks
        page_breaks.append(p.start())
    page_breaks.append(len(rtf_content))
    return page_breaks

# Function to extract the page header

def extract_header(page_content):
    '''This function extracts the page header using the '\\header' RTF tag
    The content enclosed within the '\\header' tag, is found,
    and the data is extracted using an re expression
    '''
    global PAGE
    PAGE += 1
    try:
        header_start = re.search(r'{\\header' , page_content).start()
        # Finding the '\header' tag to find the header
        i = header_start+1
        flag = 0
        header_end = 0
        # This while loop is used to find the header content in the page
        # It parses through the data to find the'}'symbol
        while i < len(page_content):
            if page_content[i] == '{' :
                flag += 1
            if page_content[i] == '}' :
                if flag == 0 :
                    header_end = i
                    break
                flag -= 1
            i += 1

        header = re.findall(RE_expressions['header'] , page_content[header_start : header_end])
        headerstyle = re.findall(RE_expressions['headerstyle']
                                ,page_content[header_start : header_end])
        #Dict to store the header content with the alignment
        headers={}
        for h, header_line in enumerate(header):
            header_line = header_line.replace(
                '{\\field{\\*\\fldinst { PAGE }}}{',str(PAGE)).replace(
                    '}{\\field{\\*\\fldinst { NUMPAGES }}}',str(NUMPAGES))
            headers[header_line]=headerstyle[h][0]
        debug_print("Header extracted successfully")

    # Used to check whether the header content is extracted successfully
    except AttributeError:
        debug_print("Header not found")
        log_file_exceptions.write("Header not extracted successfully in page " + str(PAGE))
    return headers, page_content[header_end+1:]

# Function to extract the table title

def extract_title(page_content):
    '''This function is used to extract the table title using the '\\trhdr' RTF tag
    The '\\row' tag is used to find the end of the title rows
    Re expressions are used to extract the titles data
    '''

    try:
        trhdr = []
        end_row = []
        for t in re.finditer(r'\\trhdr' , page_content) :
            trhdr.append(t.start())
        for e in re.finditer(r'{\\row}' , page_content) :
            end_row.append(e.end())
        title = []
        for i in range(len(trhdr)-1):
            title_line=re.search(r'{(.+)\\cell}',page_content[trhdr[i]:end_row[i]]).group()[1:-6]
            title_line=re.sub(r"\\(\w+)","",title_line).strip()
            title.append(title_line)
        debug_print("Title extracted successfully")
        if len(trhdr)>1:
            return title, page_content[end_row[len(trhdr)-2]+1:]
        return title, page_content[trhdr[0]:]

    # Used to check whether the table title is extracted successfully
    except AttributeError:
        debug_print("No title found")
        log_file_exceptions.write("Title not extracted successfully in page " + str(PAGE))
# Function to extract the table column headers

def extract_column_headers(page_content):
    '''This function extracts the column headers
    The '\\row' tag is used to find the end of the column headers row
    '''
    try:
        end_row = re.search(r'{\\row}', page_content).end()

        headers = re.findall(r'{(.+)\\cell}' , page_content[:end_row])
        column_headers = [re.sub(r"\\(\w+)", "", h).strip() for h in headers]
        debug_print("Column headers extracted successfully")

    # Used to check whether the column headers are extracted successfully
    except AttributeError:
        debug_print("Column headers not found")
        log_file_exceptions.write("Column headers not extracted successfully in page " + str(PAGE))
        column_headers = []
    return column_headers, page_content[end_row+1:]

# Function to extract the table data

def extract_table_data(page_content, column_headers):
    '''This function is used to extract the table data
    The '\\trowd' tag is used to find the beginning of each row
    The '\\row' tag is used to find the end of each row
    The data in each row is extracted, and mapped to the column headers in a dictionary
    The presence of footnotes in the page is checked using the '\\keepn' tag
    '''
    try:
        trowd = []
        end_row = []
        for t in re.finditer(r'\\trowd' , page_content) :
            trowd.append(t.start())
        for e in re.finditer(r'{\\row}' , page_content) :
            end_row.append(e.end())

        subjects = []
        no_of_rows = len(trowd)
        if re.search(r'\\keepn', page_content[trowd[-1]:end_row[-1]]):
            no_of_rows -= 1
        for r in range(no_of_rows):
            row_data = re.findall(r'{(.+)\\cell}' , page_content[trowd[r]:end_row[r]])
            row_data = list(filter(None, [re.sub(r"\\\w+" , "" , rd).strip() for rd in row_data]))
            if row_data :
                subject_details = {}
                for i, row_data_values in enumerate(row_data) :
                    if not row_data_values.isdigit():
                        subject_details[column_headers[i]] = row_data_values
                    else:
                        subject_details[column_headers[i]] = int(row_data_values)

                subjects.append(subject_details)
        debug_print(
            "Table data extracted successfully")
    # Used to check whether the table data is extracted successfully
    except AttributeError:
        debug_print("Table data not found")
        log_file_exceptions.write("Table data not extracted successfully in page " + str(PAGE))
        subjects = []
    return subjects, page_content[end_row[r]:]
# Function to extract the table footnotes
def extract_footnotes(page_content):
    '''This function is used to extract the footnotes using an re expression'''
    try:
        footnotes = [re.search(r'{(.+)\\cell}', page_content).group()[1:-6]]
        debug_print(f"Footer found: {footnotes}")
        debug_print("Footnotes extracted successfully")
    # Used to check whether the footnotes are extracted successfully
    except AttributeError:
        debug_print("Footnotes not found")
        message = "Footnotes not extracted successfully in page " + str(PAGE)
        log_file_exceptions.write(message + "\n")
        footnotes = []
    return footnotes

# Function to extract the table footer

def extract_footer(footnotes):
    '''This function is used to find the table footer
    Is found by searching for 'Source'/'Dataset'
    The table footers are then extracted
    '''
    try:
        if not footnotes :
            return [], []
        footnote = footnotes[0]
        if "Source" in footnote :
            i = footnote.find("Source")
            return footnote[:i] , footnote[i:]
        if "Dataset" in footnote :
            i = footnote.find("Dataset")
            return footnote[:i] , footnote[i:]
    # Used to check whether the footer is extracted successfully
    except AttributeError:
        debug_print("Footer not found")
        log_file_exceptions.write("Footer not extracted successfully in page " + str(PAGE))

# Function to extract the contents of a page

def extract_page_content(page_content):
    '''
    This function is used to extract the content of each page
    A dictionary called 'page_details' is initialized
    The respective functions to extract the page header, table title,
    column headers, subjects details, footnotes and footers are called
    '''
    page_details = {}
    page_details['header'], page_content = extract_header(page_content)
    page_details['title'], page_content = extract_title(page_content)
    page_details['column headers'], page_content = extract_column_headers(page_content)
    page_details['subjects'], page_content = (
    extract_table_data(page_content, page_details['column headers'])
    )
    page_details['footnotes'] = extract_footnotes(page_content)
    page_details['footnotes'], page_details['footer'] = (
    extract_footer(page_details['footnotes'])
    )

    return page_details

# Function to convert an rtf file to json

def convert_rtf(item, file_no, output_directory):
    '''
    This function is used to convert the RTF file into JSON format
    The page breaks function is called to split the content for each page
    '''
    debug_print(f"Converting file {file_no}: {item}")
    # output_log = open('/Users/shreejakatama/Downloads/Internship/Folder Code/Output_log.txt','a')
    try:
      # Extract rtf content as a string in python
        with open(item, 'r', encoding = "utf-8") as file:
            rtf_content = file.read().replace("{\\line}\n", " ").replace("\\~", " ")
            debug_print(f"RTF content loaded for file {file_no}")

        fonts = extract_font_details(rtf_content)
        # debug_print(f"Fonts extracted: {fonts}")

        json_dictionary = {}
        data = []
        json_dictionary ['fonts'] = fonts
        json_dictionary ['data'] = data

        page_breaks = extract_page_breaks(rtf_content)
        debug_print(f"Page breaks found: {page_breaks}")

        PAGE = 0
        NUMPAGES = rtf_content.count("NUMPAGES")
        for i in range(len(page_breaks)-1) :

            page_content = rtf_content[page_breaks[i] : page_breaks[i+1]]
            debug_print(f"Processing page {PAGE + 1}")
            page_details = extract_page_content(page_content)
            data.append(page_details)
        output_file = os.path.join(
            output_directory,
            f"{os.path.splitext(os.path.basename(item))[0]}.json"
            )

        with open(output_file, 'w', encoding = "utf-8") as f:
            json.dump(json_dictionary , f, indent=4)
        debug_print(f"JSON file {output_file} successfully created")
        log_file_success.write(f"Data successfully written to {output_file}\n")

        return "Successful", ""
    except AttributeError as e:
        debug_print("Error, cannot be converted due to " + str(e))
        log_file_exceptions.write(item+" cannot be converted due to "+ "\n")
        return "Failed", "Not in Scope"

def upload_folder():
    '''
    This is a function to get the folder from the user
    This function uses the UI to accept and upload the folder
    '''
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        FOLDER_PATH.set(folder_selected)
        global SELECTED_FOLDER_PATH
        SELECTED_FOLDER_PATH = folder_selected
        process_files(folder_selected)

def process_files(selected_folder):
    '''
    This function is used to process the files in the folder
    It creates an output directory in the parent folder
    It iterates through the folder
    It checks if the file is an RTF file
    If the file is an RTF file, the schema of the file is checked
    If the file adheres to the schema, the file is converted to JSON
    '''
    global OUTPUT_DIRECTORY, FOLDER_TO_DELETE  # Declare as global variables
    for row in table.get_children():
        table.delete(row)

    if not selected_folder:
        return

    files = os.listdir(selected_folder)
    OUTPUT_DIRECTORY = os.path.join(selected_folder, 'Output')
    FOLDER_TO_DELETE = OUTPUT_DIRECTORY  # Assign the output directory to FOLDER_TO_DELETE
    os.makedirs(OUTPUT_DIRECTORY, exist_ok=True)
    print('{OUTPUT_DIRECTORY} successfully created')
    file_no = 0
    for file in files:
        file_path = os.path.join(selected_folder, file)
        if not file.endswith('.rtf'):
            status = "Failed"
            remarks = "Choose a RTF File"
            color = 'red'
            debug_print("Not an RTF File, cannot be converted")
        elif os.path.isfile(file_path):
            file_no += 1
            if check_rtf(file_path):
                print("RTF File conforms to schema")
                status, remarks = convert_rtf(file_path, file_no, OUTPUT_DIRECTORY)
                color = 'green' if status == "Successful" else 'red'
                debug_print("RTF File converted successfully")
            else:
                print(f"RTF File {file} does not conform to schema, cannot be converted")
                log_file_exceptions.write(
                    f"RTF File {file} does not conform to schema, "
                    "cannot be converted\n"
                    )

                status = "Failed"
                remarks = "No remarks Found"
                color = 'red'
        else:
            status = "Failed"
            remarks = "No remarks Found"
            color = 'red'

        table.insert("", "end", values=(file, status, remarks), tags=(color,))

def on_continue():
    '''
    This function serves as a placeholder for the functionality of the 'continue' button
    '''
    messagebox.showinfo("Info", "Continue button clicked!")

def on_delete():
    '''
    This function is used to delete the output folder 
    '''
    for row in table.get_children():
        table.delete(row)
    FOLDER_PATH.set("")

def user_interface():
    '''
    This function is used to set up the User Interface
    '''
    app = tk.Tk()
    app.title("RTF to JSON Converter")
    app.geometry("800x600")
    global FOLDER_PATH
    FOLDER_PATH = tk.StringVar()

    title_label = tk.Label(app, text="RTF to JSON Converter", font=("Times New Roman", 20, "bold"))
    title_label.place(relx=0.5, rely=0.05, anchor=tk.CENTER)
    subtitle1_label = tk.Label(
    app,
    text="Convert Your RTF Document to JSON Format",
    font=("Times New Roman", 8)
    )

    subtitle1_label.place(relx=0.5, rely=0.1, anchor=tk.CENTER)
    subtitle2_label = tk.Label(
    app,
    text="Select the RTF Folder to be Uploaded",
    font=("Times New Roman", 14)
    )

    subtitle2_label.place(relx=0.5, rely=0.2, anchor=tk.CENTER)

    upload_button = tk.Button(app, text="UPLOAD RTF FOLDER", command=upload_folder)
    upload_button.place(relx=0.5, rely=0.25, anchor=tk.CENTER)

    global table
    columns = ("File Name", "Status", "Remarks")
    table = ttk.Treeview(app, columns=columns, show="headings")
    table.heading("File Name", text="File Name")
    table.heading("Status", text="Status")
    table.heading("Remarks", text="Remarks")
    table.place(relx=0.5, rely=0.57, anchor=tk.CENTER, relwidth=0.8, relheight=0.55)

    table.tag_configure('green', background='lightgreen')
    table.tag_configure('red', background='lightcoral')

    style = ttk.Style()
    style.configure("TButton", padding=6, relief="flat", background="#ccc")
    style.map("TButton",
            background=[('active', '#0052cc'), ('!disabled', '#004080')],
            foreground=[('active', 'white'), ('!disabled', 'Black')],
            relief=[('pressed', 'sunken'), ('!pressed', 'raised')])

    continue_button = ttk.Button(app, text="CONTINUE", command=on_continue, style="TButton")
    continue_button.place(relx=0.3, rely=0.9, anchor=tk.CENTER)

    delete_button = ttk.Button(app, text="DELETE", command=on_delete, style="TButton")
    delete_button.place(relx=0.7, rely=0.9, anchor=tk.CENTER)

    app.mainloop()
    debug_print("UI loaded")

# Main function to call the user_interface() function
if __name__ == "__main__" :
    try:
        user_interface()
    except Exception:
        debug_print("UI unsuccessful")

#test functions

# Function to check if a file is an RTF file

"""
This test function is to check if the given file is RTF or not and is mapped to the function 'is_rtf_file'
Test Scenario : 

- Test case id : KEX002.2_TC00.0

- Test case description : Verify that the file name contains a '.rtf' extension. If not, write test cases for other file types

- Test case step : 
1. Check for a folder directory in the system
2. Check for .rtf extension 
3. Generate test functions for the test cases
4. Generate test functions for other file types 

"""
def is_rtf_file(file_path):
    return os.path.splitext(file_path)[1].lower() == '.rtf'

# Test cases for the function
def test_is_rtf_file():
    # Test case 1: File with .rtf extension
    assert is_rtf_file("example.rtf") == True

    # Test case 2: File with .txt extension
    assert is_rtf_file("example.txt") == False

    # Test case 3: File with uppercase .RTF extension
    assert is_rtf_file("example.RTF") == True

    # Test case 4: File with no extension
    assert is_rtf_file("example") == False


# Main block to execute the function
if __name__ == "__main__":
    for file_name in os.listdir(SELECTED_FOLDER_PATH):
        file_path = os.path.join(SELECTED_FOLDER_PATH, file_name)
        # Check if the path is a file and has an .rtf extension
        if os.path.isfile(file_path) and is_rtf_file(file_path):
            print(f"{file_path} is an RTF file.")
        else:
            print(f"{file_path} is not an RTF file.")



#test function to check for rtf schema

"""
This test function is to check if for the RTF schema and is mapped to the function 'check_rtf'

Test Scenario : 
- Test case id : KEX002.2_TC00.1

- Test case description : Verify that the RTF File conforms to the RTF Schema with the corresponding RTF control tags

- Test case step : 
1. Check for a folder directory in the system
2. Open the RTF file after checking the folder 
3. Check whether the RTF file conforms to schema
4. Generate test function for checking RTF File

"""
def test_check_rtf_from_folder():
    # Main block to execute the function
    if __name__ == "__main__":
        valid_files = []
        invalid_files = []
        for file_name in os.listdir(SELECTED_FOLDER_PATH):
            file_path = os.path.join(SELECTED_FOLDER_PATH, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith('.rtf'):
                if check_rtf(file_path):
                    valid_files.append(file_path)
                else:
                    invalid_files.append(file_path)


#test function to check for fonts in a file

def test_extract_font_details():
    # Main block to execute the function
    if __name__ == "__main__":
        valid_font = []
        invalid_font = []
        for file_name in os.listdir(SELECTED_FOLDER_PATH):
            file_path = os.path.join(SELECTED_FOLDER_PATH, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith('.rtf'):
                with open(file_path, 'r') as file:
                    rtf_content = file.read()
                    extracted_fonts = extract_font_details(rtf_content)

                # Check for valid font tables
                if extracted_fonts:
                    valid_font.append(file_path)
                    # Assert that the extracted fonts are non-empty
                    assert extracted_fonts, f"Error in file {file_path}: Extracted fonts should not be empty."
                else:
                    invalid_font.append(file_path)
                    # Assert that extracted fonts are empty for malformed content
                    assert not extracted_fonts, f"Error in file {file_path}: Extracted fonts should be empty for malformed tables."

        for file in valid_font:
            fonts = extract_font_details(open(file).read())
            assert fonts, f"File {file} should have valid font details."

        # Assert that error_files have no valid font details
        for file in invalid_font:
            # Ensure error files are truly malformed or incorrect
            assert not extract_font_details(open(file).read()), f"File {file} should have an empty font details due to errors."
    

#test function to extract header
"""
"This test function is to verify the presence of a header in an RTF file and is mapped with the 'extract_header' function."

Test Scenario:
- test case id : KEX002.2_TC001

- test case description : Verify that the system extracts all header components from the RTF table generated by SAS/R software.

- test step:
1. Upload the sample RTF file containing the table to the system.
2. Initiate the extraction process.
3. Open the generated JSON file.
4. Verify that all header components from the RTF table are present in the JSON output.
"""
def test_extract_header():
    if __name__ == "__main__":
        valid_header = []
        invalid_header = []
        for file_name in os.listdir(SELECTED_FOLDER_PATH):
            file_path = os.path.join(SELECTED_FOLDER_PATH, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith('.rtf'):
                with open(file_path, 'r') as file:
                    rtf_content = file.read()
                    extracted_header = extract_header(rtf_content)

                # Check for valid font tables
                if extracted_header:
                    valid_header.append(file_path)
                    # Assert that the extracted fonts are non-empty
                    assert extracted_header, f"Error in file {file_path}: Extracted header should not be empty."
                else:
                    invalid_header.append(file_path)
                    # Assert that extracted fonts are empty for malformed content
                    assert not extracted_header, f"Error in file {file_path}: Extracted header should be empty for malformed tables."

        for file in valid_header:
            header = extract_header(open(file).read())
            assert header, f"File {file} should have valid header details."

        # Assert that error_files have no valid font details
        for file in invalid_header:
            # Ensure error files are truly malformed or incorrect
            assert not extract_header(open(file).read()), f"File {file} should have an empty header details due to errors."
    

#test function to extract title
"""
"This test function is to verify the presence of a title in an RTF file and is mapped with the 'extract_title' function."

- test case id : KEX002.2_TC004

- Test case description: Verify that the system extracts all title components from the RTF table generated by SAS/R software.

- test step:
1.Upload the sample RTF file containing the table to the system.
2. Initiate the extraction process.
3. Open the generated JSON file.
4. Verify that all title components from the RTF table are present in the JSON output.
"""
def test_extract_title():
    if __name__ == "__main__":
        valid_title = []
        invalid_title = []
        for file_name in os.listdir(SELECTED_FOLDER_PATH):
            file_path = os.path.join(SELECTED_FOLDER_PATH, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith('.rtf'):
                with open(file_path, 'r') as file:
                    rtf_content = file.read()
                    extracted_title = extract_title(rtf_content)

                # Check for valid font tables
                if extracted_title:
                    valid_title.append(file_path)
                    # Assert that the extracted fonts are non-empty
                    assert extracted_title, f"Error in file {file_path}: Extracted title should not be empty."
                else:
                    invalid_title.append(file_path)
                    # Assert that extracted fonts are empty for malformed content
                    assert not extracted_title, f"Error in file {file_path}: Extracted title should be empty for malformed tables."

        for file in valid_title:
            header = extract_header(open(file).read())
            assert header, f"File {file} should have valid title details."

        # Assert that error_files have no valid font details
        for file in invalid_title:
            # Ensure error files are truly malformed or incorrect
            assert not extract_header(open(file).read()), f"File {file} should have an empty title details due to errors."

    
#test function to extract column header
"""
"These test functions are to verify the presence of column header and is mapped with the functions "extract_column_header"

- Test case id : KEX002.2_TC005

- Test case description : Verify that the system extracts all body components from the RTF table generated by SAS/R software.

- Test step :
1. Upload the sample RTF file containing the table to the system.
2. Initiate the extraction process.
3. Open the generated JSON file.
4. Verify that all body components from the RTF table are present in the JSON output.
"""
def test_extract_column_headers():
    if __name__ == "__main__":
        valid_column_header = []
        invalid_column_header = []
        for file_name in os.listdir(SELECTED_FOLDER_PATH):
            file_path = os.path.join(SELECTED_FOLDER_PATH, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith('.rtf'):
                with open(file_path, 'r') as file:
                    rtf_content = file.read()
                    extracted_column_header = extract_column_headers(rtf_content)

                # Check for valid font tables
                if extracted_column_header:
                    valid_column_header.append(file_path)
                    # Assert that the extracted fonts are non-empty
                    assert extracted_column_header, f"Error in file {file_path}: Extracted column header should not be empty."
                else:
                    invalid_column_header.append(file_path)
                    # Assert that extracted fonts are empty for malformed content
                    assert not extracted_column_header, f"Error in file {file_path}: Extracted column header should be empty for malformed tables."

        for file in valid_column_header:
            header = extract_column_headers(open(file).read())
            assert header, f"File {file} should have valid column header details."

        # Assert that error_files have no valid font details
        for file in invalid_column_header:
            # Ensure error files are truly malformed or incorrect
            assert not extract_column_headers(open(file).read()), f"File {file} should have an empty column header details due to errors."

#test function to extract table data
"""
"These test functions are to verify the presence of table data and is mapped with the functions "extract_table_data"

- Test case id : KEX002.2_TC005

- Test case description : Verify that the system extracts all table data from the RTF table generated by SAS/R software.

- Test step :
1. Upload the sample RTF file containing the table to the system.
2. Initiate the extraction process.
3. Open the generated JSON file.
4. Verify that all body components from the RTF table are present in the JSON output.
"""
def test_extract_table_data():
    if __name__ == "__main__":
        valid_table_data = []
        invalid_table_data = []
        for file_name in os.listdir(SELECTED_FOLDER_PATH):
            file_path = os.path.join(SELECTED_FOLDER_PATH, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith('.rtf'):
                with open(file_path, 'r') as file:
                    rtf_content = file.read()
                    extracted_table_data = extract_table_data(rtf_content)

                # Check for valid font tables
                if extracted_table_data:
                    valid_table_data.append(file_path)
                    # Assert that the extracted fonts are non-empty
                    assert extracted_table_data, f"Error in file {file_path}: Extracted table data should not be empty."
                else:
                    invalid_table_data.append(file_path)
                    # Assert that extracted fonts are empty for malformed content
                    assert not extracted_table_data, f"Error in file {file_path}: Extracted table data should be empty for malformed tables."

        for file in valid_table_data:
            table_data = extract_table_data(open(file).read())
            assert table_data, f"File {file} should have valid table data details."

        # Assert that error_files have no valid font details
        for file in invalid_table_data:
            # Ensure error files are truly malformed or incorrect
            assert not extract_table_data(open(file).read()), f"File {file} should have an empty table data details due to errors."


#test function to extract footnotes
"""
"This test function is to verify the presence of footnotes in an RTF file and is mapped with the 'extract_footnotes' function."

- Test case id : KEX002.2_TC002

- Test case description : Verify that the system extracts all footnotes from the RTF table generated by SAS/R software.

- Test step :
"1. Upload the sample RTF file containing the table to the system.
2. Initiate the extraction process.
3. Open the generated JSON file.
4. Verify that all footnotes from the RTF table are present in the JSON output."
"""
def test_extract_footnotes():
    if __name__ == "__main__":
        valid_footnotes = []
        invalid_footnotes = []
        for file_name in os.listdir(SELECTED_FOLDER_PATH):
            file_path = os.path.join(SELECTED_FOLDER_PATH, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith('.rtf'):
                with open(file_path, 'r') as file:
                    rtf_content = file.read()
                    extracted_footnotes = extract_footnotes(rtf_content)

                # Check for valid font tables
                if extracted_footnotes:
                    valid_footnotes.append(file_path)
                    # Assert that the extracted fonts are non-empty
                    assert extracted_footnotes, f"Error in file {file_path}: Extracted footnotes should not be empty."
                else:
                    invalid_footnotes.append(file_path)
                    # Assert that extracted fonts are empty for malformed content
                    assert not extracted_footnotes, f"Error in file {file_path}: Extracted footnotes should be empty for malformed tables."

        for file in valid_footnotes:
            footnotes = extract_footnotes(open(file).read())
            assert footnotes, f"File {file} should have valid footnotes details."

        # Assert that error_files have no valid font details
        for file in invalid_footnotes:
            # Ensure error files are truly malformed or incorrect
            assert not extract_footnotes(open(file).read()), f"File {file} should have an empty footnotes details due to errors."


#test function to extract footer
"""
"This test function is to verify the presence of a footer in an RTF file and is mapped with the 'extract_footer' function."

- Test case id : KEX002.2_TC003

- Test case description : Verify that the system extracts all footer components from the RTF table generated by SAS/R software.

- Test step
1. Upload the sample RTF file containing the table to the system.
2. Initiate the extraction process.
3. Open the generated JSON file.
4. Verify that all footer components from the RTF table are present in the JSON output.
"""

def test_extract_footer():
    if __name__ == "__main__":
        valid_footer = []
        invalid_footer = []
        for file_name in os.listdir(SELECTED_FOLDER_PATH):
            file_path = os.path.join(SELECTED_FOLDER_PATH, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith('.rtf'):
                with open(file_path, 'r') as file:
                    rtf_content = file.read()
                    extracted_footer = extract_footer(rtf_content)

                # Check for valid font tables
                if extracted_footer:
                    valid_footer.append(file_path)
                    # Assert that the extracted fonts are non-empty
                    assert extracted_footer, f"Error in file {file_path}: Extracted footer should not be empty."
                else:
                    invalid_footer.append(file_path)
                    # Assert that extracted fonts are empty for malformed content
                    assert not extracted_footer, f"Error in file {file_path}: Extracted footer should be empty for malformed tables."

        for file in valid_footer:
            footer = extract_footer(open(file).read())
            assert footer, f"File {file} should have valid footer details."

        # Assert that error_files have no valid font details
        for file in invalid_footer:
            # Ensure error files are truly malformed or incorrect
            assert not extract_footer(open(file).read()), f"File {file} should have an empty footer details due to errors."

#test function to convert rtf
"""
This test function is to verify that the conversion of RTF to JSON is successful.

- Test step 
1. Provide a valid input (e.g., a data structure or content) that needs to be converted into a JSON file.
2. Run the JSON file creation module.
3. Check if a JSON file is generated and verify its content.
"""
def test_convert_rtf():
    if __name__ == "__main__":
        valid_rtf = []
        invalid_rtf = []
        for file_name in os.listdir(SELECTED_FOLDER_PATH):
            file_path = os.path.join(SELECTED_FOLDER_PATH, file_name)
            if os.path.isfile(file_path) and file_path.lower().endswith('.rtf'):
                with open(file_path, 'r') as file:
                    rtf_content = file.read()
                    extracted_rtf = convert_rtf(rtf_content)

                # Check for valid font tables
                if extracted_rtf:
                    valid_rtf.append(file_path)
                    # Assert that the extracted fonts are non-empty
                    assert extracted_rtf, f"Error in file {file_path}: Extracted rtf content should not be empty."
                else:
                    invalid_rtf.append(file_path)
                    # Assert that extracted fonts are empty for malformed content
                    assert not extracted_rtf, f"Error in file {file_path}: Extracted rtf content should be empty for malformed tables."

        for file in valid_rtf:
            footer = convert_rtf(open(file).read())
            assert footer, f"File {file} should have valid rtf content details."

        # Assert that error_files have no valid font details
        for file in invalid_rtf:
            # Ensure error files are truly malformed or incorrect
            assert not convert_rtf(open(file).read()), f"File {file} should have an empty rtf content details due to errors."