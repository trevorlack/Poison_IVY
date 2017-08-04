import win32com.client
from datetime import date, timedelta
import os
import pytz
import datetime
import holidays

import sys
from win32com.client import constants

email_dater = date.today()
print(email_dater)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
message = messages.GetFirst()

'''Wrap this in a Try Statement'''

while message:
    if message.Subject == 'Preliminary Cash Available Report':
        dater = message.CreationTime
        dater = dater.date()
        if dater == email_dater:
            email_flows = message.Body
            file = open("email_body.txt", "w")
            file.write(email_flows)
            file.close()
        break
    message = messages.GetNext ()

'''Add code to pull in the shareholder equity number to compare assets in model'''

email_flows = email_flows.splitlines()
print(email_flows)
HY_location = email_flows.index('921')
HY_ShareHolder_Trades = email_flows[HY_location+4]
HY_ShareHolder_Trades = HY_ShareHolder_Trades.strip()
HY_ShareHolder_Trades = HY_ShareHolder_Trades.replace(',','')
HY_ShareHolder_Float = float(HY_ShareHolder_Trades)
print(email_flows[HY_location+2])

IG_location = email_flows.index('924')
IG_ShareHolder_Trades = email_flows[IG_location+4]
IG_ShareHolder_Trades = IG_ShareHolder_Trades.strip()
IG_ShareHolder_Trades = IG_ShareHolder_Trades.replace(',','')
IG_ShareHolder_Float = float(IG_ShareHolder_Trades)
print(email_flows[IG_location+2])

os.chdir("R:\Fixed Income\IVY\Trading Model")
x1 = win32com.client.DispatchEx("Excel.Application")
x1.Visible = True
x1.Workbooks.Open("C:/blp/API/Office Tools/BloombergUI.xla")
x1.Workbooks.Open(os.path.join(os.getcwd(), "IVY MASTER_PROOF.xlsm"))
wb = x1.Workbooks("IVY MASTER_PROOF.xlsm")
ws = wb.Worksheets("Proof")
ws.Range("D4").Value = HY_ShareHolder_Float
ws.Range("D5").Value = IG_ShareHolder_Float
x1.ActiveWorkbook.Save()

IVY_email_dater = datetime.datetime.now(pytz.timezone('US/Eastern'))- datetime.timedelta(days = 2)
fmt = '%Y-%m-%d'
IVY_email_dater = IVY_email_dater.strftime(fmt)
#email_dater = datetime.datetime.now() - datetime.timedelta(days = 1)
print(IVY_email_dater)
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
folder = outlook.Folders('Portfolio')
inbox = folder.Folders('Inbox')

messages = inbox.Items

message = messages.GetFirst ()

while message:
    if message.Subject == 'Ivy ProShares Reports - Fixed Income (Funds 921 & 924)':
        NAV_dater = message.CreationTime
        NAV_dater = NAV_dater.strftime(fmt)
        print(NAV_dater)
        if NAV_dater > IVY_email_dater:
            attachments = message.Attachments
            for i in range(attachments.Count):
                attachment = attachments.Item(i + 1)
                print(attachment.FileName)
                attachment.SaveASFile('R:\\Fixed Income\\IVY\\Archive\\NAV Report\\' + attachment.FileName)
            break
    message = messages.GetNext ()
#x1.ActiveWorkbook.Close(True)

'''
current_hyhg_data = int(pull_hy_match_set())
print(current_hyhg_data)

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6).Folders('Scripting')

messages = inbox.Items

message = messages.GetFirst ()

while message:
    print(message.Subject[0:15])
    if message.Subject == 'Ivy ProShares Reports - Fixed Income (Funds 921 & 924)':
        File_Date = int(message.Subject[-16:-7])
        if File_Date > current_hyhg_data:
            attachments = message.Attachments
            for i in range(attachments.Count):
                attachment = attachments.Item(i+1)
                print(attachment.FileName)
                attachment.SaveASFile('R:\\Fixed Income\\IVY\\Archive\\NAV Report\\' + attachment.FileName)
    message = messages.GetNext ()
'''
