import openpyxl


def analysis(system, india_xlsx_path, notes_from_sap): 
       
    notes_from_india = set()

    ### exc handling in GUI
    workbook_india = openpyxl.load_workbook(india_xlsx_path) #wb from india BUT FORMATTED
    
    # finding the exact name
    for sheet in workbook_india.sheetnames:
        if system in sheet:
            worksheet_india = workbook_india[sheet]
            break
        # only reached if break never occured
        raise IndexError(f"\n\nSheet '{system}' could not be found in workbook '{india_xlsx_path}'")



    # get notes from India XSLX, save to set 'notes_from_india'
    for e in range (2, worksheet_india.max_row + 1):
        cell_address = 'A' + str(e)
        cell_value = worksheet_india[cell_address].value
        if cell_value is not None:
            notes_from_india.add(cell_value)

    return notes_from_sap ^ notes_from_india
