#metapub package: https://pypi.org/project/metapub/
import os
import xlwings as xw


def check_file_exist(lib_directory):

    from mainfunctions import process_ids
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

    wb = xw.books.active
    ws = xw.sheets.active
    rng = wb.app.selection

    #settings
    isfileColor = (188, 219, 255) #sky-blue
    # isfileTextColor = (41,60,94)

    totalfilecount = 0
    isfilecount = 0
    nofilecount = 0
    
    for i in list(rng):
        b = i.value
        print(b)
        if is_none(b):
            continue
        b=num_to_str(b)
        print(b)
        if "|" in b:
            continue
        isfilecurrentcell = True
        
        if len(process_ids(b,lib_directory)[4])>0:
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
