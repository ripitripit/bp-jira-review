import pandas as pd
import numpy as np
from tkinter import *
from tkinter import filedialog
import xlsxwriter

#convert report to a csv file
def toExcel(filename):
    pd.read_excel(filename, index=False).to_excel('report.xlsx', encoding='utf-8', columns=['Issue key',
    'Custom field (Client)',
    'Summary',
    'Created',
    #'Resolved',
    'Status',
    #'Custom field (Time to first response)',
    #'Custom field (Time to resolution)',
    'Reporter',
    'Assignee'])

#initialize node mapping dictionary
nodeMapping = {
"CYAAC": "7-Eleven",
"CYPWBS124": "7-Eleven",
"CYPWBS125": "7-Eleven",
"CYPWBS132": "7-Eleven",
"CYPAPP26": "7-Eleven",
"CYAAC1WBSENC02": "7-Eleven",
"CYPMSMQ01": "7-Eleven",
"CY1AAWBS01": "7-Eleven",
"CY1AAWBS02": "7-Eleven",
"CY1AAWBS03": "7-Eleven",
"CY1AAWBS04": "7-Eleven",
"CY1AAWBS05": "7-Eleven",
"CY1AAWBS06": "7-Eleven",
"CY1AAWBS07": "7-Eleven",
"CY1AAWBS08": "7-Eleven",
"CYPAPP21": "American Eagle",
"CYPWBS2017": "American Eagle",
"CYPWBS207": "American Eagle",
"CYPWBS208": "American Eagle",
"CYPWBS209": "American Eagle",
"CYPWBS106": "BP Group",
"CY2CINFMAN02": "BP Group",
"CYPWBS142": "Chevron",
"CYPAPP212": "Essilor",
"CYPWBS69": "Express",
"CYPWBS70": "Express",
"CYPWBS88": "Express",
"CYPWBS89": "Express",
"CYPWBS90": "Express",
"CYPWBS91": "Express",
"CYPWBS14": "Gamestop",
"CYPWBS16": "Gamestop",
"CYPWBS17": "Gamestop",
"CYPWBS18": "Gamestop",
"CYPWBS19": "Gamestop",
"CYPWBS45": "Gamestop",
"CYPWBS46": "Gamestop",
"CYPWBS47": "Gamestop",
"CYPWBS48": "Gamestop",
"CYPWBS80": "Hard Rock",
"CYPAPP10": "Hard Rock",
"CYPAPP221": "Hertz",
"CYPAPP222": "Hertz",
"CYPWBS65": "Kelloogs",
"CYPWBS66": "Kelloogs",
"CYPWBS67": "Kelloogs",
"CYPWBS110": "Moneygram",
"CYPWBS216": "Pep Boys",
"CYPWBS218": "Pep Boys",
"CYPWBS219": "Pep Boys",
"AWSTYOPWBS01": "Phillip Morris Japan",
"AWSTYOPWBS02": "Phillip Morris Japan",
"CY2CAUT01": "Qentelli",
"CYPWBS232": "The Children's Place",
"CYPWBS233": "The Children's Place",
"CYPWBS234": "The Children's Place",
"CYPWBS235": "The Children's Place"}

#initialize node requirement dictionary
nodeRequirement = (
"System - Active Thread",
"System - CPU Above Normal",
"System - Error Patterns Found",
"System - High Processing Time",
"System - Kong Status",
"System - Lack of Free Memory",
"System - Lack of Free Swap Space",
"System - Low Free Disk Space",
"System - Mulipath Errors Found",
"System - Oracle unable to connect",
"System - Zabbix Unreachable",
"System - Server Too Busy",
"System - Slow LDAP lookups",
"System - W3SVC is down")

#get list of summary formats from the "JIRA Ticketing Formats" SOP
tickFormatDf = pd.read_excel('Formats_092219.xlsx', encoding='utf-16', columns=["Summary Format"])
tickFormatDf["Summary Format"] = tickFormatDf["Summary Format"].str.split("-").str[:2].str.join("-")
summaryFormats = tickFormatDf["Summary Format"]

#get dictionary of client:server from "Client Name Node Mapping" SOP
##df2 = pd.read_excel('BP - Client Name Node Mapping2.xlsx', encoding='utf-8', columns=['Client', 'Associated Nodes'])
##df2 = df2.drop(['Client Name', 'Notes'], axis=1)
##df2.to_dict()

#create dataframe
def createDF():
    df = pd.read_excel('report.xlsx', encoding='utf-8')
    df['Summary'] = df['Summary'].astype(str)

    #create new columns & assume defaults are false
    df['Summary Verification'] = "False"
    df['Node Requirement'] = "False"
    df['Client Verification'] = "False"

    #does the ticket match summary per SOP?
    for x in summaryFormats:
        df.loc[df['Summary'].str.contains(x), 'Summary Verification'] = 'True'
    #does the ticket require a server node per SOP?
    for x in nodeRequirement:
        df.loc[df['Summary'].str.contains(x), 'Node Requirement'] = 'True'
    #does the ticket match the client per SOP?
    for x in nodeMapping:
        df.loc[df['Summary'].str.contains(x), 'Client Verification'] = 'True'

    #create range variables for conditional formatting
    rowCt = len(df.index)
    range1 = ('E2:E' +  str(rowCt+1))
    print('Cell range 1: ', range1)
    range2 = ('D2:D' +  str(rowCt+1))
    print('Cell range 2: ', range2)
    
    #create ExcelWriter workbook
    writer = pd.ExcelWriter('report.xlsx', engine='xlsxwriter')
    workbook  = writer.book
    formatY = workbook.add_format({'bg_color': 'yellow'})
    df.to_excel(writer, sheet_name='Sheet1')
    worksheet = writer.sheets['Sheet1']
    worksheet.conditional_format(range1,
                                     {'type': 'formula',
                                     'criteria': '=IF(J1="True",FALSE,TRUE)',
                                     'format':formatY})
    worksheet.conditional_format(range2,
                                     {'type': 'formula',
                                      'criteria': '=IF((AND(K1="True",L1="False")),TRUE,FALSE)',
                                      'format':formatY})
    writer.save()

def main():
    #create window object
    window = Tk()
    window.title('BP Reporter')
    #create run command for lambda (if needed)
    def run(command):
        (str(command))
    #create string variable for directory of file
    filepath = StringVar()
    #define buttons
    b1 = Button(window, text='Browse:', width=15, height=2, bg='white', fg='black', command=lambda: filepath.set(filedialog.askopenfilename(filetypes=(('BP Jira Report', '.xlsx'), ('All files', '*.*')))))
    b1.grid(row=0, column=0)
    b1 = Button(window, text='Convert to CSV:', width=15, height=2, bg='white', fg='black', command=lambda: run(toExcel(filepath.get())))
    b1.grid(row=1, column=0)
    b1 = Button(window, text='Run Report:', width=15, height=2, bg='white', fg='black', command=lambda: run(createDF()))
    b1.grid(row=2, column=0)
    #define labels
    l1 = Label(window, textvariable=filepath)
    l1.grid(row=0, column=1)
    #create window mainloop
    window.mainloop()
main()
