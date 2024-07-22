import json
import re
import os

# Global debugging flag
DEBUG = True

def debug_print(message):
    if DEBUG:
        print(message)

# Function to extract font details from RTF content
def extract_font_details(rtf_content):
    debug_print("Extracting font details")

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
        debug_print(f"Found font: ID = {font_id}, Name = {font_name}")
    
    return fonts


def convert_rtf(item,file_no):
    debug_print(f"Converting file {file_no}: {item}")

    # Extract rtf content as a string in python
    with open(directory+'/'+item, 'r') as file:
        rtf_content = file.read().replace("{\\line}\n", " ").replace("\\~", " ")
        debug_print(f"RTF content loaded for file {file_no}")

    fonts = extract_font_details(rtf_content)
    debug_print(f"Fonts extracted: {fonts}")

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
    debug_print(f"Page breaks found: {page_breaks}")
    
    for i in range(len(page_breaks)-1) :
        patient_dict = {}

        page_content = rtf_content[page_breaks[i] : page_breaks[i+1]]
        PAGE += 1
        debug_print(f"Processing page {PAGE}")

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
        for h in range (len(patient_dict['header'])):
            patient_dict['header'][h] = patient_dict['header'][h].replace('{\\field{\\*\\fldinst { PAGE }}}{',str(PAGE)).replace('}{\\field{\\*\\fldinst { NUMPAGES }}}',str(NUMPAGES))
        debug_print(f"Extracted header: {patient_dict['header']}")

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
        debug_print(f"Titles found: {patient_dict['title']}")

        headers = re.findall(r'{(.+)\\cell}' , page_content[trhdr[-1]:end_row[len(trhdr)-1]])
        patient_dict['headers'] = [re.sub(r"\\(\w+)", "", h).strip() for h in headers]
        debug_print(f"Headers found: {patient_dict['headers']}")

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
        debug_print(f"Footer found: {patient_dict['footer']}")

        data.append(patient_dict)
    
    output_file = newdir + '/' + item[:-4] + '.json'
    with open(output_file, 'w') as f:
        json.dump(dict1, f, indent=4)
    debug_print(f"Data successfully written to {output_file}")

    # f = open(newdir+'/'+item[:-4]+'.json', 'w')
    # json.dump(dict1, f, indent=4)
    # print("Data Successfully written for file: ",file_no)


directory = '/Users/adithi/Desktop/jenkins_folder_2'
contents = os.listdir(directory)
newdir = '/Users/adithi/Desktop/jenkins_output_folder_2'
os.mkdir(newdir)
file_no = 0
for item in contents:
    if item == '.DS_Store':
        continue
    else:
        file_no+=1
        convert_rtf(item,file_no)
