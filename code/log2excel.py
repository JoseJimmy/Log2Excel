# Initial Version of Pandas Code
import tkinter,xlsxwriter
from tkinter import filedialog,messagebox
import pandas as pd

########################################################### Get Filename from user

tkinter.Tk().withdraw()
LogFilename = filedialog.askopenfilename(title = "Select log file from PTAF ",filetypes = (("txt files","*.txt"),("all files","*.*")))
ExcelFilename = LogFilename.replace('.txt','_report.xlsx')

########################################################### Initialise Variables
MacroName,StepType,StepNo ,Macro= '--','Main',' ',False
MacroStepNo = '-'
FailsInAim = []
########################################################### Read File in list logfile

with open(LogFilename) as f:
    logfile = f.readlines()

########################################################### Iterate through log and make failure summary
for line in logfile:
    line=line.rstrip('\n')
    if ('Log started at' in line):
        date = line.split(' ')[-2]

    if('testname =' in line):
        AimName=line.split('=')[1].split('\n')[0].replace(' ', '_')

    if ('---- Step:' in line):
        temp = line.split('----')[1].strip()
        if (StepType == 'Main'):
            StepNo = temp
        else:
            MacroStepNo = temp

    if('start Macro' in line):
        Macro = True
        MacroName = line.split('worksheet')[1].split('\n')[0]
        StepType = 'Macro'

    if('end Macro' in line):
        Macro = False
        MacroName = '--'
        StepType = 'Main'
        MacroStepNo = '-'

    if ('**FAIL' in line):
        Outcome = 'Fail'
        comment = line.split('**FAIL')[1].strip()

    if ('PASS' in line):
        Outcome = 'Pass'
        comment = line.split('PASS')[1].strip()

    if ('**ERROR' in line):
        Outcome = 'Error'
        comment = line.split('**ERROR')[1].strip()

    if (('FAIL' in line) or ('PASS' in line) or ('ERROR' in line)):
        timestamp = line.split(' ')[0]
        FailsInAim.append([date+' '+timestamp,AimName,MacroName,StepNo,MacroStepNo,StepType,Outcome,comment])

header = ['Datetime','AimName','MacroName','StepNo','MacroStepNo','StepType','Outcome','comment']
x= pd.DataFrame(FailsInAim,columns = header)
x['Datetime'] = pd.to_datetime(x['Datetime'])
x.to_csv('fails.csv',index = False)



