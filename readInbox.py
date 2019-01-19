import pythoncom

import win32com.client

# Connect to Outlook, which has to be running
try:
    print "trying to open up Outlook..."
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
except:
    print "Could not open Outlook"
print "successfully connected to Outlook\n"

inbox = outlook.GetDefaultFolder(6)

#Iterate over sub-folders of inboxto find the desired folder

fldr_iterator = inbox.Folders   
desired_folder = None
while 1:
    f = fldr_iterator.GetNext()
    if not f: break
    if f.Name == 'Nim':
        print 'found "test" dir'
        desired_folder = f
        break

print desired_folder

#Retrieve messages from a folder

messages = desired_folder.Items
count = messages.count
print count

message = messages.GetFirst()
while message:
	print message.Subject
	message = messages.GetNext()

#subject = message.Subject
#print subject
