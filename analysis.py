import openpyxl
import pandas as pd

from pathlib import Path

def analysis(system, kw, year):
    #KW = input("What is the KW?: ")
    
    insap = []
    inindia = []


    ### exc handling in GUI
    # wb = openpyxl.load_workbook(os.environ['USERPROFILE'] + r'\Downloads\SAP Hinweise KW ' + str(KW) + '_' + str(year) + '.xlsx') #wb from india BUT FORMATTED
    wb = openpyxl.load_workbook(Path.home().joinpath(f"Downloads/SAP Hinweise KW {kw}_{year}.xlsx")) #wb from india BUT FORMATTED

    # lol
    if system == 'DE & CH & AT':
        try:
            ws = wb['SAP Hinweise DE & CH & AT']
        except:
            ws = wb['SAP Hinweise DE & CH & AT ']

    elif system == 'ESS':
        try:
            ws = wb['ESS']
        except:
            ws = wb['ESS ']

    elif system == 'Reisemanagement':
        try:
            ws = wb['Reisemanagement']
        except:
            ws = wb['Reisemanagement ']

    elif system == 'Successfactors':
        try:
            ws = wb['Successfactors']
        except:
            ws = wb['Successfactors ']


    #change csv into xlsx
    # data_csv = pd.read_csv(os.environ['USERPROFILE']+r'\Downloads\data.csv', sep=';', lineterminator='\n')
    # gfg = pd.ExcelWriter(os.environ['USERPROFILE']+r'\Downloads\data.xlsx')
    data_csv = pd.read_csv(Path.home().joinpath('Downloads/data.csv'), sep=';', lineterminator='\n')
    gfg = pd.ExcelWriter(Path.home().joinpath("Downloads/data.xlsx"))
    data_csv.to_excel(gfg, index = False)
    gfg.close()
    wb2 = openpyxl.load_workbook(Path.home().joinpath("Downloads/data.xlsx"))# WB from SAP
    ws2 = wb2['Sheet1']

    #put all SAP NOTES IDS from wb SAP into array 'insap'
    for i in range(2, ws2.max_row + 1):
        cellsap = 'B' + str(i)
        sapvalue = ws2[cellsap].value
        insap.append(sapvalue)

    #put all SAP NOTES from wb INDIA into array 'in india'
    for e in range (2, ws.max_row + 1):
        cellindia = 'A' + str(e)
        cellvalueindia = ws[cellindia].value
        if cellvalueindia is not None:
            inindia.append(cellvalueindia)

    #function to compare both arrays
    def Diff(insap, inindia):
        return list(set(insap) - set(inindia)) + list(set(inindia) - set(insap))
 
    return Diff(insap, inindia)

    # if len(diff_sap_notes) == 0:
    #     input('\nFINISHED! No differing entries.')
        
    # else:
    #     print('\nDiffering entries:\n')
        
    #     input('\nTo open all notes in browser, press a key and log in')
    #     webbrowser.open("https://launchpad.support.sap.com")
        
        
    #     input('\nPress any key to continue')
    #     for i in diff_sap_notes:
    #         webbrowser.open("https://launchpad.support.sap.com/#/notes/"+str(i))
