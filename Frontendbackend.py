
import json
import re
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import shutil

# Global debugging flag
DEBUG = True
output_directory = ""
folder_to_delete = ""

def debug_print(message):
    if DEBUG:
        print(message)

# Function to extract font details from RTF content
def extract_font_details(rtf_content):
    # debug_print("Extracting font details")

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
        # debug_print(f"Found font: ID = {font_id}, Name = {font_name}")
    
    return fonts


def convert_rtf(item,file_no, output_directory):
    debug_print(f"Converting file {file_no}: {item}")

    try:
        # Extract rtf content as a string in python
        with open(item, 'r') as file:
            rtf_content = file.read().replace("{\\line}\n", " ").replace("\\~", " ")
            debug_print(f"RTF content loaded for file {file_no}")

        fonts = extract_font_details(rtf_content)
        # debug_print(f"Fonts extracted: {fonts}")

        dict1 = {}
        data = []
        dict1['fonts'] = fonts
        dict1['data'] = data

        NUMPAGES = rtf_content.count("NUMPAGES")
        PAGE = 0

        page_breaks = []
        for p in re.finditer(r"\\endnhere", rtf_content):
            page_breaks.append(p.start())
        page_breaks.append(len(rtf_content))
        # debug_print(f"Page breaks found: {page_breaks}")
        
        for i in range(len(page_breaks)-1) :
            patient_dict = {}

            page_content = rtf_content[page_breaks[i] : page_breaks[i+1]]
            PAGE += 1
            # debug_print(f"Processing page {PAGE}")

            header_start = re.search(r'{\\header' , page_content).start()
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

            patient_dict['header'] = re.findall(r'{(?!\\)(.+)\\cell}' , page_content[header_start : header_end])
            # print(page_content[header_start : header_end])
            for h in range (len(patient_dict['header'])):
                patient_dict['header'][h] = patient_dict['header'][h].replace('{\\field{\\*\\fldinst { PAGE }}}{',str(PAGE)).replace('}{\\field{\\*\\fldinst { NUMPAGES }}}',str(NUMPAGES))
            # debug_print(f"Extracted header: {patient_dict['header']}")

            page_content = page_content[header_end+1:]

            trhdr = []
            trowd = []
            end_row = []
            for t in re.finditer(r'\\trhdr' , page_content) :
                trhdr.append(t.start())
            for t in re.finditer(r'\\trowd' , page_content) :
                trowd.append(t.start())
            for e in re.finditer(r'{\\row}' , page_content) :
                end_row.append(e.end())
            
            patient_dict['title'] = []
            for i in range(len(trhdr)-1):
                title = re.search(r'{(.+)\\cell}' , page_content[trhdr[i]:end_row[i]]).group()[1:-6]
                title = re.sub(r"\\(\w+)", "", title).strip()
                patient_dict['title'].append(title)
            # debug_print(f"Titles found: {patient_dict['title']}")

            headers = re.findall(r'{(.+)\\cell}' , page_content[trhdr[-1]:end_row[len(trhdr)-1]])
            patient_dict['headers'] = [re.sub(r"\\(\w+)", "", h).strip() for h in headers]
            # debug_print(f"Headers found: {patient_dict['headers']}")

            patient_dict['patients'] = []
            for r in range(len(trhdr), len(trowd)-1) :
                row_data = re.findall(r'{(.+)\\cell}' , page_content[trowd[r]:end_row[r]])
                row_data = list(filter(None, [re.sub(r"\\\w+" , "" , rd).strip() for rd in row_data]))
                if row_data :
                    patient_details = {}
                    for i in range(len(row_data)) :
                        patient_details[patient_dict['headers'][i]] = row_data[i] if not row_data[i].isdigit() else int(row_data[i])
                    patient_dict['patients'].append(patient_details)

            patient_dict['footer'] = [re.search(r'{(.+)\\cell}', page_content[trowd[-1]:end_row[-1]]).group()[1:-6]]
            # debug_print(f"Footer found: {patient_dict['footer']}")

            data.append(patient_dict)
        
        output_file = os.path.join(output_directory, f"{os.path.splitext(os.path.basename(item))[0]}.json")
        with open(output_file, 'w') as f:
            json.dump(dict1, f, indent=4)
        debug_print(f"Data successfully written to {output_file}")
        # output_log.write(f"Data successfully written to {output_file}\n")
        return "Successful", ""
    
    except Exception as e:
        debug_print("Error converting")
        return "Failed", "Corrupt file detected"

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
    file_no = 0
    for file in files:
        file_path = os.path.join(folder_path, file)
        if not file.endswith('.rtf'):
            status = "Failed"
            remarks = "Choose a RTF File"
            color = 'red'
        elif os.path.isfile(file_path):
            file_no += 1
            status, remarks = convert_rtf(file_path, file_no, output_directory)
            color = 'green' if status == "Successful" else 'red'
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

app = tk.Tk()
app.title("RTF to JSON Converter")
app.geometry("800x600")

folder_path = tk.StringVar()

title_label = tk.Label(app, text="RTF to JSON Converter", font=("Arial", 20, "bold"))
title_label.place(relx=0.5, rely=0.05, anchor=tk.CENTER)

subtitle1_label = tk.Label(app, text="Convert Your RTF Document to JSON Format", font=("Arial", 8))
subtitle1_label.place(relx=0.5, rely=0.1, anchor=tk.CENTER)

subtitle2_label = tk.Label(app, text="Select the RTF Folder to be Uploaded", font=("Arial", 14))
subtitle2_label.place(relx=0.5, rely=0.2, anchor=tk.CENTER)

upload_button = tk.Button(app, text="UPLOAD RTF FOLDER", command=upload_folder)
upload_button.place(relx=0.5, rely=0.25, anchor=tk.CENTER)

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
