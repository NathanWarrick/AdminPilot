import win32com.client
import os
import re
import PyPDF2


outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6) # "6" refers to the index of a folder - in this case the Inbox

folders = {}    

# This funcion assigns all folders indexes to a dictionary
# To retrieve a index use folders['Folder Name']
for f in range(100):
    try:
        fdr = inbox.Folders(f)
        name = fdr.Name
        folders.update({name: f})
        print(f, name)
    except:
        pass


                
for msg in inbox.Folders(folders['1 Excursions']).Items:
    if re.match('Excursion Approved - .+', msg.Subject) and msg.FlagStatus == 2: # Grab flagged excursion approvals
        print('Subject: ' + msg.Subject)

        match = 0
        i = 1
        while match == 0:
            if not re.search('Medical Form', str(msg.Attachments.Item(i))) and not re.search('.png', str(msg.Attachments.Item(i))): # Check that the attachment doesn't contain the terms medical form or .png in the title
                attachment = msg.Attachments.Item(i)
                print('Path: ' + os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                attachment.SaveAsFile(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                match = 1
                
                # Read the PDF
                form = PyPDF2.PdfReader(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName)
                data = form.pages[0].extract_text()
                #print(data)
                #print(data.find('Date /s'))
                #print(data.find('Details of Cost'))
                
            else:
                i = i + 1  # Try the next attachment
                
        os.remove(os.getcwd() + r'\workspace\clients\email_client\downloads' + '\\' + attachment.FileName) # Remove file upon completion


