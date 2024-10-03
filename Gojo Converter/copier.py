from utilities import find_header, convert_weight

def copier(source_sheet, target_sheet, header_name, target_column, start_row, rows_to_search, enable_column, disable_columns):
    #use pos description as ref as only one with somehting in every row
    header_column, header_row = find_header(source_sheet, rows_to_search, header_name)
    ref_column, ref_row = find_header(source_sheet, rows_to_search, 'POS Description')
    source_row = header_row + 1 

    target_row = start_row
    ref_row += 1  

    #iterate over current column and reference column to skip any rows with empty name cells. zip helps account for empty cells apparently
    for ref_cell, cell in zip(source_sheet.iter_rows(min_row=ref_row, min_col=ref_column, max_col=ref_column), 
                              source_sheet.iter_rows(min_row=source_row, min_col=header_column, max_col=header_column)):

        ref_value = ref_cell[0].value 
        value = cell[0].value  

        #skip to next iteration for items with no name
        if ref_value is None:
            continue  
        
        #if weight, convert
        if header_name == 'Weight**':
            value = convert_weight(value) #convert_weight function from utilities.py

        #only write one off fields when pos description is being copied
        if header_name == 'POS Description':
            target_sheet.cell(row=target_row, column=enable_column, value='Y') #enable at location matching sheet name
            for disable_column in disable_columns:
                target_sheet.cell(row=target_row, column=disable_column, value='N') #disable at locations not matching sheet name
            
            target_sheet.cell(row=target_row, column=13, value='visible')  #visibility
            target_sheet.cell(row=target_row, column=14, value='Prepared food and beverage')  #item type
            target_sheet.cell(row=target_row, column=16, value='Y')  #delivery
            target_sheet.cell(row=target_row, column=24, value='Y')  #auto add to bill
        
        if value is not None:
            target_sheet.cell(row=target_row, column=target_column, value=value)

        target_row += 1  

