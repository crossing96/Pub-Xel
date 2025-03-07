# Pub-Xel - A Biomedical Reference Management Tool
# Copyright (C) 2024  Jongyeob Kim <jongyeobkim@pubxel.org>
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program. If not, see <https://www.gnu.org/licenses/>.

import os
import re
import xlwings as xw
import requests
from bs4 import BeautifulSoup
import pyperclip
import json
import platform

#appdata
os_name = platform.system() #"Windows" or "Darwin"
#appdata
if os_name == "Windows":
    appdatadir = os.path.join(os.getenv('APPDATA'), 'pubxel')
elif os_name == "Darwin":
    appdatadir = os.path.expanduser("~/Library/Application Support/pubxel")
os.makedirs(appdatadir, exist_ok=True)
settings_path = os.path.join(appdatadir,"settings.json")

def load_settings():
    if not os.path.exists(settings_path):
        return {}
    with open(settings_path, 'r') as file:
        return json.load(file)

def save_settings(settings):
    with open(settings_path, 'w') as file:
        json.dump(settings, file, indent=4)

def save_settings_key(settings, key, value):
    settings[key] = value
    save_settings(settings)
    return settings

def set_preserve_order(input_list):
    seen = set()
    return [x for x in input_list if not (x in seen or seen.add(x))]

def files_name_to_path(fileName, mainlibdir, seclibdir=[]): # assumes that files are definitely present
    if not fileName:
        return None
    pathList = []
    seclibdir = [d for d in seclibdir if os.path.isdir(d)]
    if not seclibdir:
        for file in fileName:
            pathList.append(os.path.join(mainlibdir, file))
    else:
        for file in fileName:
            if os.path.exists(os.path.join(mainlibdir, file)):
                pathList.append(os.path.join(mainlibdir, file))
            else:
                for directory in seclibdir:
                    if os.path.exists(os.path.join(directory, file)):
                        pathList.append(os.path.join(directory, file))
                        break
    return pathList

def string_to_list(input):
    if input is None:
        return None
    #tab and line break to delimiter
    input = input.replace("\t","|")
    input = input.replace("\r\n","|")
    input = input.replace("\n","|")
    input = input.replace("<","|")
    input = input.replace(">","|")
    input = input.replace(":","|")
    # 중복되는 | 제거
    to_remove = "|"
    pattern = "(?P<char>[" + re.escape(to_remove) + "])(?P=char)+"
    input = re.sub(pattern, r"\1", input)
    input = str.split(input,sep="|")
    # list 공백값 제거
    input = [item for item in input if item.strip() != ""]
    # apply strip()
    input = [i.strip() for i in input]
    # Remove duplicates from the list
    input = list(set_preserve_order(input))
    return input

def list_to_string(list,chr=60): #accepts string_to_list outcomes. assumes that duplicates already removed.  
    if list is None:
        return ""
    # Check if the list is empty
    if not list:
        return "" 
    listlength = len(list)
    longlist = False
    if listlength > 3:
        list = list[:3]
        longlist = True
    # Concatenate the list with each elements wrapped in <>
    string = f"{listlength} Selection{'s' if listlength > 1 else ''}: "+"<" + "> <".join(list) + ">"
    if len(string) > chr and string.count(">")>1:
        #remove the last <>
        string = string[:string.rfind(">",0,len(string)-1)+1]
        longlist = True
    if len(string) > chr and string.count(">")>1:
        #remove the last <>
        string = string[:string.rfind(">",0,len(string)-1)+1]
    if len(string)> chr or longlist:
        string = string[:chr]+" ..."
    return string

def process_ids(ids, maindir, seclibdir=[]):

    seclibdir = [d for d in seclibdir if os.path.isdir(d)]

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

    if ids is None:
        return valid_ids, pubmed_ids, non_pubmed_valid_ids, invalid_ids, valid_ids_with_m_files, valid_ids_without_m_files, pubmed_ids_with_m_files, pubmed_ids_without_m_files, pubmed_ids_with_s_files, pubmed_ids_without_s_files, all_m_files, all_s_files,nonpubmed_ids_with_m_files,nonpubmed_ids_without_m_files
    
    # string to list
    if isinstance(ids, str):
        ids = [ids]
    # Remove duplicates from the list
    ids = list(set_preserve_order(ids))

    for id in ids:
        if re.match(r"^[0-9]+$", id):
            valid_ids.append(id)
            pubmed_ids.append(id)
            m_files = []
            s_files = []

            filelist = os.listdir(maindir)
            if seclibdir:
                for directory in seclibdir:
                    filelist.extend(os.listdir(directory))

            for filename in filelist:
                if re.match(rf"^{id}[^0-9]", filename):
                    if re.match(rf"^{id}\.[^.]+$", filename):
                        m_files.append(filename)
                    elif re.match(rf"^{id}[^.]+", filename) or re.match(rf"^{id}\.[^.]+\..+$", filename):
                        s_files.append(filename)
            m_files = list(dict.fromkeys(m_files))
            s_files = list(dict.fromkeys(s_files))
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
        elif re.match(r"^[^0-9].*$", id):
            valid_ids.append(id)
            non_pubmed_valid_ids.append(id)
            m_files = []
            filelist = os.listdir(maindir)
            if seclibdir:
                for directory in seclibdir:
                    filelist.extend(os.listdir(directory))
            for filename in filelist:
                if filename.startswith(id) and re.match(rf"^{id}\.[^.]+$", filename):
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

def copy_list(lst):
    # Check if the list is empty
    if not lst:
        return None  # Don't do anything
    # Concatenate the list with a line break as a delimiter
    result = '\n'.join(lst)
    # Copy the string to the clipboard
    pyperclip.copy(result)

def trim_range(rng,reselect=True):
    wb = xw.books.active
    ws = xw.sheets.active

    used_range = ws.used_range
    rng = ws.range((rng.row, rng.column), 
                (rng.rows[-1].row, rng.columns[-1].column))
    if(rng.row < used_range.row):
        rng = ws.range((used_range.row, rng.column), 
                (rng.rows[-1].row, rng.columns[-1].column))
    if(rng.rows[-1].row > used_range.rows[-1].row):
        rng = ws.range((rng.row, rng.column), 
                (used_range.rows[-1].row, rng.columns[-1].column))
    if(rng.column < used_range.column):
        rng = ws.range((rng.row, used_range.column), 
                (rng.rows[-1].row, rng.columns[-1].column))
    if(rng.columns[-1].column > used_range.columns[-1].column):
        rng = ws.range((rng.row, rng.column), 
                (rng.rows[-1].row, used_range.columns[-1].column))

    # Get the values of the range as a list of lists
    rowmin=rng.rows[-1].row
    rowmax=rng.row
    columnmin=rng.columns[-1].column
    columnmax=rng.column

    notNone = False
    # Iterate through each cell in the range
    for row in rng.rows:
        for cell in row:
            if cell.value is None:
                continue
            notNone = True
            if cell.row < rowmin:
                rowmin = cell.row
            if cell.row > rowmax:
                rowmax = cell.row
            if cell.column < columnmin:
                columnmin = cell.column
            if cell.column > columnmax:
                columnmax = cell.column

    # check if any of the values are not None
    if not notNone:
        print("No values found in the range")
        return rng
    else:
        rng = ws.range((rowmin, columnmin), (rowmax, columnmax))
        if reselect:
            rng.select()
        return rng

def check_file_exist(mainlibdir,seclibdir=[]):
    
    def is_none(value):
        return value is None

    def num_to_str(num):
        # Check if the input is a string
        if isinstance(num, str):
            return num
        elif num % 1 == 0:  # Check if the number is a whole number
            return str(int(num))  # Convert to integer before converting to string to remove the decimal part
        else:
            return str(num)  # If not a whole number, convert to string directly

    try:
        app = xw.apps.active
        wb = xw.books.active
        ws = xw.sheets.active
    except:
        raise ValueError("Please open the Excel Worksheet first.")

    try:
        rng = wb.app.selection
    except:
        raise ValueError("No selection made. Please make a selection in the Excel sheet.")

    if rng is None:
        raise ValueError("No selection made.. Please make a selection in the Excel sheet.")

    if rng.count > 1000:
        raise ValueError("Please select 1000 or fewer cells!")
    
    rng = trim_range(rng,reselect=True)

    if rng is None:
        raise ValueError("No selection made.. Please make a selection in the Excel sheet.")


    # Step 1: Identify column interval of the selection
    column_interval = (rng.column, rng.column + rng.columns.count - 1)
    # Step 2: Select a range of the first row of the sheet, ranging from the column range of the selection
    header_range = ws.range((1, column_interval[0]), (1, column_interval[1]))

    # Check if any range is selected for header_range
    if header_range.count == 0:
        header_range.select()
        raise ValueError("Error: Please ensure the following before trying again.:\n1. The table header must be located in the first row of the entire Excel sheet (Row 1).\n2. The column header containing PMIDs is labeled as 'Ref'.")

    # Check if "ref" is present in the header_range
    ref_count = sum(1 for cell in header_range if cell.value is not None and cell.value.lower() == "ref")
    print("ref_count: ", ref_count)
    if ref_count == 0:
        header_range.select()
        raise ValueError("Error: Please ensure the following before trying again..:\n1. The table header must be located in the first row of the entire Excel sheet (Row 1).\n2. The column header containing PMIDs is labeled as 'Ref'.")
    elif ref_count > 1:
        header_range.select()
        raise ValueError("Error: Multiple 'ref' columns found. Please ensure there is only one 'ref' column in the table header.")
    
    # Check if "ref" is present in the header_range and get its location
    ref_column = None
    for cell in header_range:
        if cell.value is not None and cell.value.lower() == "ref":
            ref_column = cell.column
            break
    
    
    extended_selection = ws.range((rng.row, ref_column), (rng.row + rng.rows.count - 1, ref_column))
    extended_selection.select()
    rng = wb.app.selection

    #settings
    isfileColor = (188, 219, 255) #sky-blue

    totalfilecount = 0
    isfilecount = 0
    nofilecount = 0
    
    for i in list(rng):
        b = i.value
        if is_none(b):
            continue
        b=num_to_str(b)
        if "|" in b:
            continue
        isfilecurrentcell = True
        
        if len(process_ids(b,mainlibdir,seclibdir)[4])>0:
            isfilecount =isfilecount+1
            totalfilecount = totalfilecount+1
        else: 
            nofilecount=nofilecount+1
            totalfilecount = totalfilecount+1
            isfilecurrentcell = False
        
        if isfilecurrentcell:
            i.color= isfileColor
            # i.font.color = isfileTextColor
        elif i.color == isfileColor:
        # elif i.color == isfileColor and i.font.color == isfileTextColor:
            i.color= None
            
    print("requested cells: "+str(rng.count)+"\n"+
    "requested files: "+str(totalfilecount)+"\n"+
    "isfile count: "+str(isfilecount)+"\n"+
    "nofile count: "+str(nofilecount)+"\n")

    return (str(rng.count)+" cells checked."+"\n"+
        "Cells with files are now colored blue."
        )

def value_from_dict(dictionary, key, outputtype="string", default=""): # outputtype: "string", "first", "list"
    if outputtype not in ["string", "first", "list"]:
        raise ValueError("Invalid outputtype. Expected 'string', 'first', or 'list'.")
    if key in dictionary:
        value = dictionary[key]
        if outputtype == "string":
            return value
        elif outputtype == "first":
            return value.split('|')[0]
        elif outputtype == "list":
            return value.split('|')
    return default

def obtain_pubmed_data(PMID_list):
    if not isinstance(PMID_list, list):
        PMID_list = [PMID_list]

    PMID_list = [str(PMID) for PMID in PMID_list]
    PMID_list = [PMID.lstrip("0") for PMID in PMID_list]
    PMID_list = list(set_preserve_order(PMID_list))  # Remove duplicates
    PMID_list = [PMID for PMID in PMID_list if PMID.isnumeric()]
    PMID_list = [PMID for PMID in PMID_list if len(PMID) < 9]

    def html_to_dict(text):
        text = text.replace(" \r\n      ", " ")
        text = text.replace("\r\n      ", " ")
        result = {}
        for line in text.strip().split('\r\n'):
            if '- ' in line:
                key, value = line.split('- ', 1)
                key = key.rstrip()  # Remove trailing spaces from key
                value = value.rstrip()
                if key in result:
                    result[key] += '|' + value
                else:
                    result[key] = value
        return result
        
    def get_first_doi_value(input):
        if isinstance(input, str):
            if input.endswith(" [doi]"):
                return input[:-6]  # Remove " [doi]" from the end
            return ""  # Return "NA" if the string does not end with " [doi]"
        elif isinstance(input, list):
            for value in input:
                if value.endswith(" [doi]"):
                    return value[:-6]  # Remove " [doi]" from the end
            return ""  # Return "" if no such value is found in the list
        else:
            return ""  # Return "" if input is neither a string nor a list
        
    def NAifempty(value):
        if value == "" or value is None:
            return "NA"
        else:
            return value

    # If PMID_list has more than one component, concatenate it to a string with delimiter ","
    if len(PMID_list) > 1:
        PMID_list_str = ",".join(PMID_list)
    else:
        PMID_list_str = PMID_list[0] if PMID_list else ""

    url = "https://api.ncbi.nlm.nih.gov/lit/ctxp/v1/pubmed/?format=medline&id="+ PMID_list_str

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
    except requests.exceptions.HTTPError as e:
        raise ValueError(f"HTTP Error.\nInvalid PMID(s). Please try again.\n{e}")
    except requests.exceptions.ConnectionError:
        raise ValueError("Connection Error.\nFailed to connect. Please check your internet connection.")
    except requests.exceptions.Timeout:
        raise ValueError("Timeout Error.\nThe request timed out. Please try again later.")
    except requests.exceptions.RequestException as e:
        raise ValueError(f"Error.\nAn error occurred: {e}")
    
    html_doc = response.text
    soup = BeautifulSoup(html_doc, 'html.parser')
    data=soup.get_text()

    if not data.startswith("PMID"):
        raise ValueError("Error.\nInvalid PMID(s). Please try again.")

    # Initialize the dictionary
    html_dict = {}
    segments = data.split("\r\n\r\n")
    # Process each segment
    for segment in segments:
        lines = segment.split("\r\n")
        for line in lines:
            if line.startswith("PMID"):
                # Extract the string after "PMID - "
                pmid_key = line.split("-", 1)[1].strip()
                html_dict[pmid_key] = segment
                break  # Stop searching once we find the "PMID" line

    article_dict = {}

    #input values
    for PMID in PMID_list:
        if PMID not in html_dict:
            continue
        # Convert the text to a dictionary
        PMID_dict = html_to_dict(html_dict[PMID])
        # Kim JY
        firstauthor = value_from_dict(PMID_dict, "AU", outputtype="first")
        # Kim
        firstauthorlastname = value_from_dict(PMID_dict, "FAU", outputtype="first").split(",",1)[0]
        title = value_from_dict(PMID_dict, "TI")
        year = value_from_dict(PMID_dict, "DP").split(' ', 1)[0]
        source = value_from_dict(PMID_dict, "SO").split("doi: ", 1)[0].split("Epub ", 1)[0].split("eCollection ", 1)[0].split("PMID: ", 1)[0].rstrip(" ")
        #firstauthorlastnameetal: Kim, et al. or Kim or GBD Collaborators
        if len(value_from_dict(PMID_dict, "AU", outputtype="list")) >= 2:
            firstauthorlastnameetal = firstauthorlastname+" et al."
        elif len(value_from_dict(PMID_dict, "AU", outputtype="list")) == 1:
            firstauthorlastnameetal = firstauthorlastname
        elif value_from_dict(PMID_dict,"CN"):
            firstauthorlastnameetal = value_from_dict(PMID_dict,"CN",outputtype="first")
        else:
            firstauthorlastnameetal = ""
        #authoryear: Kim, et al., 2009 or Kim, 2009 or GBD Collaborators, 2011
        authoryear = firstauthorlastnameetal + ", " + year
        cite = authoryear+ "." + "\n"+ title + "\n" + source
        cite_maincheckbox = title + "\n" + source

        article_info = {
            "PMID": value_from_dict(PMID_dict, "PMID"),
            "title": title,
            "abstract": value_from_dict(PMID_dict, "AB"),
            "journal": value_from_dict(PMID_dict, "TA"),
            "authors": value_from_dict(PMID_dict, "AU"),
            "firstauthor": firstauthor,
            "firstauthorlastnameetal" : firstauthorlastnameetal,
            "authoryear" : authoryear,
            "date": value_from_dict(PMID_dict, "DP"),
            "year": year,
            "doi": get_first_doi_value(value_from_dict(PMID_dict, "AID")),
            "source" : source,
            "cite": cite,
            "cite_maincheckbox": cite_maincheckbox,
            "link": f"https://pubmed.ncbi.nlm.nih.gov/{PMID}"
        }
        article_dict[PMID] = article_info
    
    return article_dict

def input_pubmed_data():
    #settings
    #header {column name : requested varaible}. column name all lower case
    header = {"ref":"pmid","doi":"doi",
    "firstauthor":"fa","author":"au","year":"yr","authoryear":"fayr",
    "journal":"jo","title":"ti","abstract":"ab","citation":"cite","output2":"ou2","authors":"au2",
    "if2022":"if2022","citation2022":"cite2022",
    "if2023":"if2023","citation2023":"cite2023"}

    header2 = {}
    try:
        app = xw.apps.active
        wb = xw.books.active
        ws = xw.sheets.active
    except:
        raise ValueError("Please open the Excel Worksheet first.")

    try:
        rng = wb.app.selection
    except:
        raise ValueError("No selection made. Please make a selection in the Excel sheet.")
    
    # rng = trim_range(rng,reselect=True): this action is already done in check_file_exist

    if rng is None:
        raise ValueError("No selection made.. Please make a selection in the Excel sheet.")
    
    if rng.count > 200:
        raise ValueError("Please select 200 or fewer cells!")
    
    ###Extend the selection to the entire row
    ###identify relevant first row selection, and expand rng
    # Step 1: Identify column interval of the selection
    column_interval = (rng.column, rng.column + rng.columns.count - 1)
    # Step 2: Select a range of the first row of the sheet, ranging from the column range of the selection
    header_range = ws.range((1, column_interval[0]), (1, column_interval[1]))
    # Step 3: Extend the range to include the entire row
    header_range = header_range.expand('right')

    # Check if any range is selected for header_range
    if header_range.count == 0:
        header_range.select()
        raise ValueError("Error: Please ensure the following before trying again.:\n1. The table header must be located in the first row of the entire Excel sheet (Row 1).\n2. The column header containing PMIDs is labeled as 'Ref'.")

    # Check if "ref" is present in the header_range
    ref_count = sum(1 for cell in header_range if cell.value is not None and cell.value.lower() == "ref")
    print("ref_count: ", ref_count)
    if ref_count == 0:
        header_range.select()
        raise ValueError("Error: Please ensure the following before trying again..:\n1. The table header must be located in the first row of the entire Excel sheet (Row 1).\n2. The column header containing PMIDs is labeled as 'Ref'.")
    elif ref_count > 1:
        header_range.select()
        raise ValueError("Error: Multiple 'ref' columns found. Please ensure there is only one 'ref' column in the table header.")

    # Step 4: Return the column interval of the header_range
    new_column_interval = (header_range.column, header_range.column + header_range.columns.count - 1)
    
    # Step 5: Extend the initial selection to include the columns identified in new_column_interval, then select the header_range
    extended_selection = ws.range((rng.row, new_column_interval[0]), (rng.row + rng.rows.count - 1, new_column_interval[1]))
    extended_selection.select()
    rng = wb.app.selection
        
        
    def NAifempty(value):
        if value == "" or value is None:
            return "NA"
        else:
            return value

    #identify existent headers
    for i in list(rng[0,:]):
        if ws[0,i.column-1].value is None:
            continue
        if ws[0,i.column-1].value.lower() in header:
            header2[header.get(ws[0,i.column-1].value.lower())]=i.column-1

    rng_col_range = range(rng[:,0].row-1, rng[:,0].row-1+len(rng[:,0]))
    requested_ref_count = sum(1 for number in rng_col_range if number != 0) # length of rng_col_range but minus 1 if contain 0. 

    #impact factor
    #impact factor 파일 업데이트 하면 맨앞 맨끝 문자 정리좀 해주셈
    # Get the directory of the current script
    script_dir = os.path.dirname(os.path.abspath(__file__))
    data_dir = os.path.join(script_dir, 'data')

    impactfactor2022 = {}
    impactfactor2023 = {}

    #IF2022
    importIF2022 = header2.get("if2022",-1)>=0 or header2.get("cite2022",-1)>=0
    if importIF2022:
        IF2022_path = os.path.join(data_dir, 'impactfactor2022.txt')
        print("load IF2022")
        with open(IF2022_path, "r", encoding="utf8") as file:
            lines = file.readlines()
            for line in lines[1:]:  # Skip the header line
                parts = line.strip().split('\t')
                if len(parts) == 3:
                    journal, IF2022, quartile2022 = parts
                    impactfactor2022[journal] = (IF2022, quartile2022)

    #IF2023
    importIF2023 = header2.get("if2023",-1)>=0 or header2.get("cite2023",-1)>=0
    if importIF2023:
        IF2023_path = os.path.join(data_dir, 'impactfactor2023.txt')
        print("load IF2023")
        with open(IF2023_path, "r", encoding="utf8") as file:
            lines = file.readlines()
            for line in lines[1:]:  # Skip the header line
                parts = line.strip().split('\t')
                if len(parts) == 3:
                    journal, IF2023, quartile2023 = parts
                    impactfactor2023[journal] = (IF2023, quartile2023)

    PMID_list = []
    identifiedPMID_list=[]
    unidentifiedPMID_list=[]
    nonPMID_list = []

    print(header2)
    # Get all PMIDs possible by i's and make them to a list
    PMIDs = [ws[i, header2.get("pmid")].value for i in rng_col_range]

    for PMID in PMIDs:
        if isinstance(PMID, (int, float)) and PMID > 0 and PMID.is_integer():
            PMIDstring = str(int(PMID))
            PMID_list.append(str(int(PMIDstring)))
        elif isinstance(PMID, str) and PMID.replace('.', '', 1).isdigit() and float(PMID).is_integer() and float(PMID) > 0:
            PMIDstring = str(int(float(PMID)))
            PMID_list.append(str(int(PMIDstring)))
        else:
            nonPMID_list.append(PMID)
    
    PMID_dicts = obtain_pubmed_data(PMID_list)

    if not isinstance(PMID_dicts, dict):
        return

    #input values
    for i in rng_col_range:
        
        if(i==0):
            continue

        PMID = ws[i,header2.get("pmid")].value

        # Check if PMID is a natural number or a float that represents a natural number
        if isinstance(PMID, (int, float)) and PMID > 0 and PMID.is_integer():
            PMIDstring = str(int(PMID))
        elif isinstance(PMID, str) and PMID.replace('.', '', 1).isdigit() and float(PMID).is_integer() and float(PMID) > 0:
            PMIDstring = str(int(float(PMID)))
        else:
            continue

        if PMIDstring not in PMID_dicts:
            unidentifiedPMID_list.append(PMIDstring)
            continue
        else:
            identifiedPMID_list.append(PMIDstring)
        
        if PMID_dicts.get(PMIDstring,""):
            PMID_dict = PMID_dicts[PMIDstring] 
        else:
            continue

        journal = value_from_dict(PMID_dict,"journal","string","")
        title = value_from_dict(PMID_dict,"title","string","")
        source = value_from_dict(PMID_dict,"source","string","")
        firstauthorlastnameetal = value_from_dict(PMID_dict,"firstauthorlastnameetal","string","")
        cite = firstauthorlastnameetal.rstrip(".")+ ". "+ title + " " + source

        if importIF2022:
            try:
                IF2022 = impactfactor2022.get(journal.upper())[0]
            except Exception:
                IF2022=""
            if header2.get("cite2022",-1)>=0:
                # source: Create a pattern with a capturing group, and then use re.sub with the pattern and replacement
                pattern = re.escape(journal)
                replacement = r"\1 (IF: " + IF2022 + ")"
                source2022 = re.sub(f"({pattern})", replacement, source, count=1)
                cite2022 = firstauthorlastnameetal.rstrip(".")+ ". "+ title + " " + source2022

        if importIF2023:
            try:
                IF2023 = impactfactor2023.get(journal.upper())[0]
            except Exception:
                IF2023=""
            if header2.get("cite2023",-1)>=0:
                pattern = re.escape(journal)
                replacement = r"\1 (IF: " + IF2023 + ")"
                source2023 = re.sub(f"({pattern})", replacement, source, count=1)
                cite2023 = firstauthorlastnameetal.rstrip(".")+ ". "+ title + " " + source2023

        
        if header2.get("doi",-1)>=0:
            ws[i,header2.get("doi")].value = value_from_dict(PMID_dict,"doi","string","NA")
        if header2.get("au2",-1)>=0:
            ws[i,header2.get("au2")].value = value_from_dict(PMID_dict,"authors","string","NA")
        if header2.get("au",-1)>=0:
            ws[i,header2.get("au")].value = value_from_dict(PMID_dict,"authors","string","NA")
        if header2.get("fa",-1)>=0:
            ws[i,header2.get("fa")].value = NAifempty(firstauthorlastnameetal)
        if header2.get("ti",-1)>=0:
            ws[i,header2.get("ti")].value = NAifempty(title)
        if header2.get("ab",-1)>=0:
            ws[i,header2.get("ab")].value = value_from_dict(PMID_dict,"abstract","string","NA")
        if header2.get("jo",-1)>=0:
            ws[i,header2.get("jo")].value = value_from_dict(PMID_dict,"journal","string","NA")
        if header2.get("yr",-1)>=0:
            ws[i,header2.get("yr")].value = value_from_dict(PMID_dict,"year","string","NA")
        if header2.get("fayr",-1)>=0:
            ws[i,header2.get("fayr")].value = value_from_dict(PMID_dict,"authoryear","string","NA")
        if header2.get("cite",-1)>=0:
            ws[i,header2.get("cite")].value = NAifempty(cite)
        if header2.get("if2022",-1)>=0:
            ws[i,header2.get("if2022")].value = NAifempty(IF2022)
        if header2.get("cite2022",-1)>=0:
            ws[i,header2.get("cite2022")].value = NAifempty(cite2022)
        if header2.get("if2023",-1)>=0:
            ws[i,header2.get("if2023")].value = NAifempty(IF2023)
        if header2.get("cite2023",-1)>=0:
            ws[i,header2.get("cite2023")].value = NAifempty(cite2023)

    # # Enable screen updating
    # app.api.ScreenUpdating = True

    if True: 
        print("Number of requested references: "+str(requested_ref_count))
        print("non-PMID references: " + str(nonPMID_list))
        print("PMIDs: " + str(PMID_list))
        print("identified PMIDs: " + str(identifiedPMID_list))
        print("Unidentified PMIDs: " + str(unidentifiedPMID_list))
    
    return "Import successful"

