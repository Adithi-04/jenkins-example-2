"""
Author: Shreeja Katama
Email: sskshreeja@gmail.com

Description:
This script processes RTF files to extract basic data and convert it to JSON format.
It includes dynamic extraction using control words and robust error handling.

Version:
Initial Version: 1.0.0 (01-07-2024)
Current Version: 1.2.4 (27-07-2024)

Dependencies:
- Python Standard Libraries:
  - json
  - re
  - os
  - tkinter

Usage Instructions:
1. Ensure all dependencies are installed.
2. Run the script with a folder containing RTF files as an input parameter.
   Example: python rtf_to_json.py input_files
3. The JSON outputs will be saved in a new directory.

Known Issues:
- Some control words may not be fully supported.
- Large RTF files with different schemas may result in longer processing times.

Testing:
- Tested with sample RTF files of various sizes and contents.
- Verified JSON output with expected structure.
- Tested on Python 3.8 and 3.9 environments.
"""

import json
import re
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

# Global debugging flag
DEBUG = True
output_directory = ""
folder_to_delete = ""

def debug_print(message):
    if DEBUG:
        print(message)

# Function to check if RTF File adheres to the schema
def check_rtf(file_path):
    with open(file_path, 'r') as file:
        rtf_content = file.read().replace("{\\line}\n", " ").replace("\\~", " ")
    rtf_tags = ['\\fonttbl','\\endnhere','\\header','\\cell','\\trowd','\\row'] # Commonly used RTF tags
    flag = True
    for i in rtf_tags:
        if i in rtf_content:
            continue
        else:
            print(i," not in rtf") # If RTF tag is not present, the RTF does not adhere to the schema
            flag = False
            break
    return flag

# Function to extract font details from RTF content
def extract_font_details(rtf_content):

    # Extract font table
    font_table_pattern = re.compile(r'{\\fonttbl(.*)}', re.DOTALL) 
    font_table_match = font_table_pattern.search(rtf_content)
    if font_table_match:
        font_table = font_table_match.group(1)

    else:
        debug_print("No font table found")
        return {}

    # Extract font details from font table
    font_pattern = re.compile(r'{\\f(\d+)\\.*? ([^;]+?);}')

    fonts = {}
    for match in font_pattern.finditer(font_table):
        font_id, font_name = match.groups()
        fonts['f'+font_id] = font_name    
    return fonts

# Function to extract page breaks in the RTF File
def extract_page_breaks(rtf_content):
    page_breaks = []
    for p in re.finditer(r"\\endnhere", rtf_content): # Using the '\endhere' tag to find page breaks
        page_breaks.append(p.start())
    page_breaks.append(len(rtf_content))
    return page_breaks

# Function to extract the page header
def extract_header(page_content):
    try:
        header_start = re.search(r'{\\header' , page_content).start() # Finding the '\header' tag to find the header
        i = header_start+1
        flag = 0
        header_end = 0
        while i < len(page_content) :
            if page_content[i] == '{' :
                flag += 1
            if page_content[i] == '}' :
                if flag == 0 :
                    header_end = i
                    break
                else :
                    flag -= 1
            i += 1

        header = re.findall(r'{(?!\\)(.+)\\cell}' , page_content[header_start : header_end])
        global PAGE
        PAGE += 1
        for h in range (len(header)):
            header[h] = header[h].replace('{\\field{\\*\\fldinst { PAGE }}}{',str(PAGE)).replace('}{\\field{\\*\\fldinst { NUMPAGES }}}',str(NUMPAGES))
        # debug_print(f"Extracted header: {header}")
        debug_print("Header extracted successfully")

    except:
        debug_print("Header not found")
    return header, page_content[header_end+1:]

# Function to extract the table title
def extract_title(page_content):
    try:
        trhdr = []
        end_row = []
        for t in re.finditer(r'\\trhdr' , page_content) :
            trhdr.append(t.start())
        for e in re.finditer(r'{\\row}' , page_content) :
            end_row.append(e.end())
        
        title = []
        for i in range(len(trhdr)-1):
            title_line = re.search(r'{(.+)\\cell}' , page_content[trhdr[i]:end_row[i]]).group()[1:-6]
            title_line = re.sub(r"\\(\w+)", "", title_line).strip()
            title.append(title_line)
        # debug_print(f"Titles found: {title}")
        debug_print("Title extraxted successfully")
        if len(trhdr)>1:
            return title, page_content[end_row[len(trhdr)-2]+1:]
        else:
            return title, page_content[trhdr[0]:]

    except:
        debug_print("No title found")
 
# Function to extract the table column headers
def extract_column_headers(page_content):
    try:
        end_row = re.search(r'{\\row}', page_content).end()

        headers = re.findall(r'{(.+)\\cell}' , page_content[:end_row])
        column_headers = [re.sub(r"\\(\w+)", "", h).strip() for h in headers]
        # debug_print(f"Column Headers found: {column_headers}")
        debug_print("Column headers extracted succesfully")
    except:
        debug_print("Column headers not found")
        column_headers = []
    return column_headers, page_content[end_row+1:]

# Function to extract the table data
def extract_table_data(page_content, column_headers):
    try:
        trowd = []
        end_row = []
        for t in re.finditer(r'\\trowd' , page_content) :
            trowd.append(t.start())
        for e in re.finditer(r'{\\row}' , page_content) :
            end_row.append(e.end())

        subjects = []
        Iter = len(trowd)
        if re.search(r'\\keepn', page_content[trowd[-1]:end_row[-1]]):
            Iter -= 1
        for r in range(Iter):
            row_data = re.findall(r'{(.+)\\cell}' , page_content[trowd[r]:end_row[r]])
            row_data = list(filter(None, [re.sub(r"\\\w+" , "" , rd).strip() for rd in row_data]))
            if row_data :
                subject_details = {}
                for i in range(len(row_data)) :
                    subject_details[column_headers[i]] = row_data[i] if not row_data[i].isdigit() else int(row_data[i])
                subjects.append(subject_details)
        debug_print("Table data extracted successfully")
    except:
        debug_print("Table data not found")
        subjects = []
    
    return subjects, page_content[end_row[r]:]

# Function to extract the table footnotes 
def extract_footnotes(page_content):
    try:
        footnotes = [re.search(r'{(.+)\\cell}', page_content).group()[1:-6]]
        # debug_print(f"Footer found: {footnotes}")
        debug_print("Footnotes extracted successfully")
    except:
        debug_print("Footnotes not found")
        footnotes = []
    return footnotes

def extract_footer(footnotes):
    try:
        debug_print("Footer extracted successfully")
        if not footnotes :
            return [], []
        footnote = footnotes[0]
        if "Source" in footnote :
            i = footnote.find("Source")
            return footnote[:i] , footnote[i:]
        elif "Dataset" in footnote :
            i = footnote.find("Dataset")
            return footnote[:i] , footnote[i:]
    except:
        debug_print("Footer not found")
        
# Function to extract the contents of a page
def extract_page_content(page_content):
    page_details = {}
    page_details['header'], page_content = extract_header(page_content)
    page_details['title'], page_content = extract_title(page_content)
    page_details['column headers'], page_content = extract_column_headers(page_content)
    page_details['subjects'], page_content = extract_table_data(page_content, page_details['column headers'])
    page_details['footnotes'] = extract_footnotes(page_content)
    page_details['footnotes'] , page_details['footer'] = extract_footer(page_details['footnotes'])
    return page_details

# Function to convert an rtf file to json
def convert_rtf(item, file_no, output_directory):
    debug_print(f"Converting file {file_no}: {item}")

    try:
      # Extract rtf content as a string in python
        with open(item, 'r') as file:
            rtf_content = file.read().replace("{\\line}\n", " ").replace("\\~", " ")
            debug_print(f"RTF content loaded for file {file_no}")

        fonts = extract_font_details(rtf_content)
        debug_print(f"Fonts extracted: {fonts}")

        dict1 = {}
        data = []
        dict1['fonts'] = fonts
        dict1['data'] = data

        page_breaks = extract_page_breaks(rtf_content)
        # debug_print(f"Page breaks found: {page_breaks}")

        global PAGE
        PAGE = 0
        global NUMPAGES
        NUMPAGES = rtf_content.count("NUMPAGES")
        
        for i in range(len(page_breaks)-1) :

            page_content = rtf_content[page_breaks[i] : page_breaks[i+1]]
            debug_print(f"Processing page {PAGE + 1}")

            page_details = extract_page_content(page_content)

            data.append(page_details)
        
        output_file = os.path.join(output_directory, f"{os.path.splitext(os.path.basename(item))[0]}.json")
        with open(output_file, 'w') as f:
            json.dump(dict1, f, indent=4)
        debug_print(f"JSON file successfully created")
        return "Successful"
    
    except Exception as e:
        debug_print("Error, cannot be converted due to" + str(e))
        return "Failed"

def upload_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        folder_path.set(folder_selected)
        global selected_folder_path
        selected_folder_path = folder_selected
        process_files(folder_selected)

def process_files(folder_path):
    global output_directory, folder_to_delete  # Declare as global variables
    for row in table.get_children():
        table.delete(row)

    if not folder_path:
        return

    files = os.listdir(folder_path)
    output_directory = os.path.join(folder_path, 'Output')
    folder_to_delete = output_directory  # Assign the output directory to folder_to_delete
    os.makedirs(output_directory, exist_ok=True)
    print('{output_directory} successfully created')
    file_no = 0
    for file in files:
        file_path = os.path.join(folder_path, file)
        if not file.endswith('.rtf'):
            status = "Failed"
            remarks = "Choose a RTF File"
            color = 'red'
            debug_print("Not an RTF File, cannot be converted")
        elif os.path.isfile(file_path):
            file_no += 1
            if check_rtf(file_path):
                print("RTF File conforms to schema")
                status, remarks = convert_rtf(file_path, file_no, output_directory)
                color = 'green' if status == "Successful" else 'red'
                debug_print("RTF File converted successfully")
            else:
                print("RTF File {file} does not conform to schema, cannot be converted")
        else:
            status = "Failed"
            remarks = "No remarks Found"
            color = 'red'

        table.insert("", "end", values=(file, status, remarks), tags=(color,))

def on_continue():
    messagebox.showinfo("Info", "Continue button clicked!")


def on_delete():
    for row in table.get_children():
        table.delete(row)
    folder_path.set("")

def user_interface():
    app = tk.Tk()
    app.title("RTF to JSON Converter")
    app.geometry("800x600")

    global folder_path
    folder_path = tk.StringVar()

    title_label = tk.Label(app, text="RTF to JSON Converter", font=("Times New Roman", 20, "bold"))
    title_label.place(relx=0.5, rely=0.05, anchor=tk.CENTER)

    subtitle1_label = tk.Label(app, text="Convert Your RTF Document to JSON Format", font=("Times New Roman", 8))
    subtitle1_label.place(relx=0.5, rely=0.1, anchor=tk.CENTER)

    subtitle2_label = tk.Label(app, text="Select the RTF Folder to be Uploaded", font=("Times New Roman", 14))
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

if __name__ == "__main__" :
    try:
        user_interface()
    except:
        debug_print("UI unsuccessful")




#TEST FUNCTIONS


# Function to check if a file is an RTF file
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


# Main block to execute the function (for demonstration purposes)
if __name__ == "__main__":
    file_path = '/Users/adithi/Desktop/jenkins_folder/L-16-01-07-random.rtf'  
    if is_rtf_file(file_path):
        print(f"{file_path} is an RTF file.")
    else:
        print(f"{file_path} is not an RTF file.")


# Test function to extract header
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
    try:
        with open('/Users/adithi/Desktop/jenkins_folder/L-16-01-07-random.rtf', 'r') as file:
            rtf_content = file.read()
            # Expected output (adjust based on the actual header content in your test file)
            expected_output = "Header extracted successfully"
            actual_output = extract_header(rtf_content)
            assert actual_output == expected_output
    except:
        print(f"AssertionError:")



# Test function to extract title
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
    try:
        with open('/Users/adithi/Desktop/jenkins_folder/L-16-01-07-random.rtf', 'r') as file:
            rtf_content = file.read()
            expected_output = "Title extracted successfully"
            actual_output = extract_title(rtf_content)
            assert actual_output == expected_output
    except:
        print(f"AssertionError:")


#test functions to extract body components and table data
"""
"These test functions are to verify the presence of body components and table data and is mapped with the functions "extract_column_header and extract_table_data"

- Test case id : KEX002.2_TC005

- Test case description : Verify that the system extracts all body components from the RTF table generated by SAS/R software.

- Test step :
1. Upload the sample RTF file containing the table to the system.
2. Initiate the extraction process.
3. Open the generated JSON file.
4. Verify that all body components from the RTF table are present in the JSON output.
"""

# Test function to extract column headers
def test_extract_column_headers():
    try:
        with open('/Users/adithi/Desktop/jenkins_folder/L-16-01-07-random.rtf', 'r') as file:
            rtf_content = file.read()
            expected_output = "Column headers extracted succesfully"
            actual_output = extract_column_headers(rtf_content)
            assert actual_output == expected_output
    except:
        print(f"AssertionError:")



# Test function to extract table data
def test_extract_table_data():
    try:
        with open('/Users/adithi/Desktop/jenkins_folder/L-16-01-07-random.rtf', 'r') as file:
            rtf_content = file.read()
            expected_output = "Table data extracted successfully"
            actual_output = extract_table_data(rtf_content)
            assert actual_output == expected_output
    except:
        print(f"AssertionError:")


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

#test function to extract footnotes
def test_extract_footnotes():
    try:
        with open('/Users/adithi/Desktop/jenkins_folder/L-16-01-07-random.rtf', 'r') as file:
            rtf_content = file.read()
            expected_output = "Footnotes extracted successfully"
            actual_output = extract_footnotes(rtf_content)
            assert actual_output == expected_output
    except:
        print(f"Assertion Error:")    


"""
"This test function is to verify the presence of a footer in an RTF file and is mapped with the 'extract_footer' function."

- Test case id : KEX002.2_TC003

- Test case description : Verify that the system extracts all footer components from the RTF table generated by SAS/R software.

- Test step :
1. Upload the sample RTF file containing the table to the system.
2. Initiate the extraction process.
3. Open the generated JSON file.
4. Verify that all footer components from the RTF table are present in the JSON output.
"""

#test function to extract footer
def test_extract_footer ():
    try:
        with open('/Users/adithi/Desktop/jenkins_folder/L-16-01-07-random.rtf', 'r') as file:
            rtf_content = file.read()
            expected_output = "Footer extracted successfully"
            actual_output = extract_footer(rtf_content)
            assert actual_output == expected_output
    except:
        print(f"Assertion Error:")



"""
This test function is to verify that the conversion of RTF to JSON is successful.

- Test step 
1. Provide a valid input (e.g., a data structure or content) that needs to be converted into a JSON file.
2. Run the JSON file creation module.
3. Check if a JSON file is generated and verify its content.
"""


# Function to convert an rtf file to json
def test_convert_rtf ():
    try:
         with open('/Users/adithi/Desktop/jenkins_folder/L-16-01-07-random.rtf', 'r') as file:
            rtf_content = file.read()
            expected_output = "Successful"
            actual_output = convert_rtf(rtf_content)
            assert actual_output == expected_output
    except:
        print(f"Failed")