import re
# removes non-numeric characters and converts to kg or litres
def convert_weight(weight_str):
    if weight_str is None:
        return None
        
    # Convert the weight to a string if it’s not already one
    if not isinstance(weight_str, str):
        weight_str = str(weight_str)
        
    try:
        weight_str = re.sub(r'[^0-9.]', '', weight_str).strip()  
        if weight_str:  
            weight_grams = float(weight_str)  
            return weight_grams / 1000  # Convert grams to kg
        else:
            print(f"No valid numeric value found in weight '{weight_str}'. Skipping.")
            return None  
    except ValueError:
        print(f"Invalid weight value '{weight_str}' for conversion. Skipping.")
        return None  


#clear square template
def init(target_sheet, reset_start_row, overwrite):
    if overwrite == 1:
        for row in target_sheet.iter_rows(min_row=reset_start_row):
            for cell in row:
                cell.value = None
        print("Contents of the target sheet have been cleared.")
    else:
        print("Contents of the target sheet will not be cleared. Appending data instead.")

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
