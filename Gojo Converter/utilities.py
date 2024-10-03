import re
# removes non-numeric characters and converts to kg or litres
def convert_weight(weight_str):
    #handle empty cell
    if weight_str is None:
        return None
        
    #handle not string
    if not isinstance(weight_str, str):
        weight_str = str(weight_str)
    
    #strip non-numeric characters and convert
    try:
        weight_str = re.sub(r'[^0-9.]', '', weight_str).strip()  
        if weight_str:  
            weight_grams = float(weight_str)  
            return weight_grams / 1000 
        #no number err
        else:
            print(f"No valid number found in weight '{weight_str}'.")
            return None  
    #debug
    except ValueError:
        print(f"Invalid weight value '{weight_str}' for conversion.")
        return None  

#clear square template
def init(target_sheet, reset_start_row, overwrite):
    if overwrite == 1:
        for row in target_sheet.iter_rows(min_row=reset_start_row):
            for cell in row:
                cell.value = None
        print("Overwriting") #debug
    else:
        print("Appending") #debug

#find header
def find_header(source_sheet, rows_to_search, header):
    source_col = None  
    header_row = None
    for row in source_sheet.iter_rows(max_row=rows_to_search):  
                for cell in row:
                    if cell.value == header:
                        source_col = cell.column
                        header_row = cell.row  
                        break
                if source_col:
                    break
    return source_col, header_row

#find next empty row in square template under headers
def find_next_empty_row(ws, start_row=1):
    for row in range(start_row, ws.max_row + 2):
        if all(cell.value is None for cell in ws[row]):
            return row
