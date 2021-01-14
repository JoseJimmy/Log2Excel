import tkinter,xlsxwriter
from tkinter import filedialog,messagebox

########################################################### Get Filename from user
tkinter.Tk().withdraw()
LogFilename = filedialog.askopenfilename(title = "Select log file from PTAF ",filetypes = (("txt files","*.txt"),("all files","*.*")))
ExcelFilename = LogFilename.replace('.txt','_report.xlsx')

########################################################### Initialise Variables
AimNames,MacroStat = [['No_Aim_defined']],{}
MacroName,StepType,StepTag = '--','Main',' '
Errors,Fails,MacroFails,MainFails,i,Macro = 0, 0, 0, 0,0,False
FailsInAim=[]
FailLog=[['LineNo','Aim','StepTag','StepType','MacroName','Log']]
AimNameFound= False

########################################################### Read File in list logfile
s = open(LogFilename, 'rt')
logfile=[]
line = s.readline().replace('\n', '')
while line != '':
    logfile.append(line)
    line = s.readline().replace('\n', '')
s.close()

########################################################### Iterate through log and make failure summary
########################################################### log of failed steps and macro fail stats

for line in logfile:

    if('testname =' in line):

        Aim=line.split('=')[1].split('\n')[0].replace(' ', '_')
        FailsInAim.append([MainFails,MacroFails,Errors,MainFails+MacroFails+Errors])
        Errors,Fails,MacroFails,MainFails = 0,0,0,0

        AimNames.append([Aim])

    if('start Macro' in line):
        Macro = True
        MacroName = aim_name=line.split('orksheet')[1].split('\n')[0]
        ParentStep = StepTag
        StepType = 'Macro'

    if('end Macro' in line):
        Macro = False
        MacroName = '--'
        StepType_prev= StepType
        StepType = 'Main'
        StepTag = ParentStep

    if ('**FAIL' in line):
        if(Macro == True):
            MacroFails+=1
            if(MacroName in MacroStat):
                MacroStat[MacroName] +=1
            else:
                MacroStat[MacroName] =1
        else:
            MainFails += 1
        FailLog.append([i,Aim,StepTag,StepType,MacroName,line])

    if ('**ERROR' in line):
        Errors += 1
        FailLog.append([i,Aim,StepTag,StepType,MacroName,line])

    if( '---- Step:' in line):
        StepTag = line.split('----')[1].strip()
        StepTag = ParentStep+' > Macro '+StepTag if(Macro) else StepTag
    i+=1
FailsInAim.append([MainFails, MacroFails, Errors, MainFails + MacroFails + Errors])
del FailsInAim[0] # delete first entry - redundant
SummaryData = [a + b for a, b in zip(AimNames, FailsInAim)] # cmbine name and data
SummaryData.insert(0,['Aim Name ', 'Failed mains','Failed macros ','Errors', 'All Fails ']) #insert header

# Log file parse and summary done
# Created SummaryData (list), FailsInAim(list) & MacroStat (dictionary)
########################################################### Open Excel files and add work sheets
# Create a workbook and  worksheets
workbook = xlsxwriter.Workbook(ExcelFilename)
worksheet = workbook.add_worksheet('Summary')
worksheet1 = workbook.add_worksheet('FailedSteps')
########################################################### Set formatting and column row sizes
# formatting options
cell_format = workbook.add_format()
cell_format.set_align('left')
cell_format.set_align('top')
bold = workbook.add_format({'bold': True})
# Another comment
# Row columns sizes
worksheet.set_column('A:A', 32)
worksheet.set_column('B:F', 15)

worksheet1.set_column('A:C', 12)
worksheet1.set_column('B:C', 35)
worksheet1.set_column('D:D', 15)
worksheet1.set_column('E:E', 35)
worksheet1.set_column('F:F', 125)

# 1. Iterate through SummaryData to write failure summry
worksheet.write_row(0, 0, SummaryData[0],bold)
for row_num, data in enumerate(SummaryData[1:]):
    worksheet.write_row(row_num+1, 0, data)

# 2. Iterate through MacroStat to write macro summry
row = len(SummaryData)+2
worksheet.write(row, 0, 'MacroName',bold)
worksheet.write(row, 1, 'Fails',bold)
for i, (k, v) in enumerate(MacroStat.items(), start=1):
    worksheet.write(row+i, 0, k)
    worksheet.write(row+i, 1, v)

# 3. Iterate through MacroStat to write failed steps log data
worksheet1.write_row(0, 0, FailLog[0],bold)
for row_num, data in enumerate(FailLog[1:]):
    worksheet1.write_row(row_num+1, 0, data,cell_format)
worksheet1.autofilter(0,0,5,len(FailLog))
worksheet1.freeze_panes(1, 0)
messagebox.showinfo("Done", "Created \n"+ExcelFilename)

########################################################### Close and cleanup
workbook.close()
globals().clear()
