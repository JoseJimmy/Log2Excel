# Initial Version of Pandas Code
import tkinter,xlsxwriter
from tkinter import filedialog,messagebox
import pandas as pd
#Log-9 wrapped log processing in a function
########################################################### Get Filename from user


def Log2Dataframe(LogFilename):
    ########################################################### Initialise Variables
    MacroName,StepType,StepNo ,Macro= '--','Main','NotFound',False
    MacroStepNo = '-'
    LogAppended = False
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
            AimName=line.split('=')[1].split('\n')[0]

        if ('---- Step:' in line):
            temp = line.split('----')[1].strip().split()
            temp = (temp[1]) if len(temp) == 2 else StepNo
            timestamp = line.split(' ')[0]

            if((LogAppended == False) and (StepNo !='NotFound')): # to add steps which does not have any variables to check / comments
                FailsInAim.append([date + ' ' + timestamp, AimName, MacroName, StepNo, MacroStepNo, StepType, '--', '--'])
            LogAppended = False

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
            FailsInAim.append([date+' '+timestamp,AimName,MacroName,StepNo,MacroStepNo,StepType,Outcome,comment])
            LogAppended = True
    header = ['Datetime','AimName','MacroName','StepNo','MacroStepNo','StepType','Outcome','comment']

    df= pd.DataFrame(FailsInAim,columns = header)
    df['Datetime'] = pd.to_datetime(df['Datetime'])
    return df


tkinter.Tk().withdraw()
LogFilename = filedialog.askopenfilename(title = "Select log file from PTAF ",filetypes = (("txt files","*.txt"),("all files","*.*")))
csvFilename = LogFilename.replace('.txt','_report.csv')



AimData = Log2Dataframe(LogFilename)
AimData.to_csv("v.csv",index = False)

#%%
#Aim	Step	REQ	Test Outcome	Test Date	Tester Name	Tester Comments	Systems/Software Analysis Date	Engineer Name	Engineer Comments	Test Result
out = pd.DataFrame() #creates a new dataframe that's empty
cols =  ['Aim Name','Step','REQ','Test Outcome','Date','Tester Name','Tester Comment']
#%%
steps = []
outcome = []
comment = []
aimid = []
no = []
blanks=[]
dates = []
out = pd.DataFrame( columns=cols)
for i,aim in enumerate(pd.unique(AimData.AimName)):
    df = AimData[AimData.AimName == aim]
    for step in pd.unique(df.StepNo):

        temp = df[df.StepNo == step]
        temp = temp.sort_values(['StepNo'])
        result = '-'
        text = ""

        if ((temp['Outcome'] == 'Pass').any()):
            result = 'Pass'
            text = ""

        if((temp['Outcome']=='Fail').any()):
            result = 'Fail'
            text = ('\n'.join((temp[temp['Outcome']=='Fail'].comment.values)))

        dates.append(max(temp['Datetime']).strftime("%d/%m/%y"))
        blanks.append("")
        aimid.append(aim)
        outcome.append(result)
        comment.append(text)
        steps.append(step)
        no.append(i+1)
    summary = pd.DataFrame(list(zip(aimid,steps,blanks,outcome,dates,blanks,comment)), columns=cols)
    out = pd.concat([out, summary], join="inner")

out.to_csv("u.csv",index=False)
#
# #%% Write the dataframe without the header and index.
# import xlsxwriter
# tkinter.Tk().withdraw()
# reportFilename = filedialog.askopenfilename(title = "Select 325 File ",filetypes = (("Excel files","*.xlsx"),("all files","*.*")))
#%%
#
# with pd.ExcelWriter(reportFilename, engine="openpyxl", mode="a") as writer:
#     df.to_excel(writer, sheet_name="Test Results", startrow=9,index = False,header= False)
# wb = UpdateWorkbook(r'C:\Users\vince\Project\Spreadsheet.xlsx', worksheet=1)
# df_2 = fd_frames['key_2']
# wb['M6:M25'] = df_2['2017'].values
# wb.save()
# #%%
#
# with pd.ExcelWriter(reportFilename, engine='openpyxl', mode="a") as ew:
#     out.to_excel(ew, startrow=9, startcol=0)