import os
os.system('cls')  
import win32com.client
import time

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

inbox = outlook.GetDefaultFolder(6)
phishing = outlook.GetDefaultFolder(6).Folders['Suspicious Email'].Folders['Phishing']
spam = outlook.GetDefaultFolder(6).Folders['Suspicious Email'].Folders['Spam']
nigelphishing = outlook.GetDefaultFolder(6).Folders['nigelphishingfolder']
nigelspam = outlook.GetDefaultFolder(6).Folders['nigelspamfolder']

messages = nigelphishing.Items
messages.Sort("[ReceivedTime]", True)
message = messages.GetFirst()
tempFolder = r"C:\test\phishing"

if os.path.isdir(tempFolder):
    print("Creating Folder")
else:
    print("The Temp Folder isn't there. Creating...")
    os.mkdir(tempFolder)
    print("Folder " + tempFolder + " created successfully ")

count = 0
for message in messages:
            if '[SUSPICIOUS EMAIL]' in (message.Subject):
                print('Handling ' + message.Subject)
                count +=1
                #time.sleep(0.5)
                for attachment in message.Attachments:
                    #attachment.FileName.encode("ascii", "ignore")
                    attachment.FileName.encode("utf-8")
                 #   attachment.SaveAsFile(tempFolder + '\\' + str(count) + ' ' + attachment.FileName)
                    attachment.SaveAsFile(tempFolder + '\\' + attachment.FileName)
                 