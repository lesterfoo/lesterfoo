#Preamble
## Relevant emails are saved into "IT Assist" folder in Outlook using rules.
## Files are saved to "C:/Users/lester.foo/Videos/ServiceNow Incident Files (FWD)". Change the directory accordingly.
## Change date accordingly

#Import modules
from datetime import timedelta
import datetime
import os
import win32com.client
from zipfile import ZipFile
import shutil
import xlwings as xw

#Define variables
date = '09Mar2022'
date_obj = datetime.datetime.strptime(date, '%d%b%Y')
date_formatted = date_obj.date()

#Create new folder with date
directory = "C:/Users/lester.foo/Videos/ServiceNow Incident Files (FWD)"
path = os.path.join(directory,date)
os.mkdir(path)

#Initiate Outlook folders
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
ia = inbox.folders("IT Assist")
messages = ia.Items

#Identify emails to download
subject_Singtel = ["[External] Singtel - Accenture Problem SLA Report", "[External] Singtel - Accenture Problem Task "
                                                                    "Report",
           "[External] Singtel - Accenture Incident SLA Report"]
subject_Optus = ["[External] Accenture Problem SLA Report",
           "[External] Accenture Problem Task Report", "[External] Accenture Incident SLA Report"]

#Download attachments into newly created folder
#Singtel
for m in list(messages):
    for s in subject_Singtel:
        if m.subject == s and m.senton.date() == date_obj.date():
            attachments = m.Attachments
            attachment = attachments.Item(1)
            attachment_name = str(attachment)
            attachment.SaveASFile(path+'/'+attachment_name)

#Optus
for m in list(messages):
    for s in subject_Optus:
        if m.subject == s and m.senton.date() == date_obj.date():
            attachments = m.Attachments
            attachment = attachments.Item(1)
            attachment_name = str(attachment)
            attachment.SaveASFile(path+'/'+s+'.zip')

#Extract Optus files from zip folders
zip_list = ['[External] Accenture Incident SLA Report.zip', '[External] Accenture Problem SLA Report.zip',
            '[External] Accenture Problem Task Report.zip']
for z in zip_list:
    # opening the zip file in READ mode
    with ZipFile(path+"/"+z, 'r') as zip:
        zip.extractall(path=path)

#Copy macro file into folder
shutil.copy(directory+'/'+'SN_Convert_inc_v0.3.xlsm', path)

#Change macro date
wb = xw.Book(path+'/'+'SN_Convert_inc_v0.3.xlsm')
sheet = wb.sheets('Summary')
date_obj_minus_1d = date_obj - timedelta(days=1)
date_input = date_obj_minus_1d.strftime('%m/%d/%Y') #m and d are flipped when copied to excel
sheet.range('J3').value = date_input + ' ' + '8:00:00 AM'
wb.save(path+'/'+'SN_Convert_inc_v0.3.xlsm')

#Run macro
xl = win32com.client.Dispatch("Excel.Application")
xl.Workbooks.Open(os.path.abspath(path+'/'+'SN_Convert_inc_v0.3.xlsm'), ReadOnly=1)
xl.Application.Run("SN_Convert_inc_v0.3.xlsm!Module1.Button1_Click")
xl.Application.Quit()
del xl