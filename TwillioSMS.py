import logging as log
from twilio.rest import Client
from openpyxl import load_workbook

#Logging Config
log.basicConfig(filename='ActivityLog.log', level=log.DEBUG, format='%(asctime)s %(levelname)-8s %(message)s', datefmt='%Y-%m-%d %H:%M:%S')

#Constants
account = ""
token = ""
client = Client(account,token)

#Variables
wb = load_workbook('data.xlsx')
ws = wb['Sheet1'] #ws is woksheet [Sheet1] in workbook wb
firstNames = []
lastNames = []
phoneNumbers = []
customFields = []
messages = []
count = 0
countOfMessages = 0
approval = []

#Data initialization
for column in ws.columns:
    for cell in column:
        if(count is 0): #Count is whichever column it is on
            firstNames.append(cell.value)
        elif(count is 1):
            lastNames.append(cell.value)
        elif (count is 2):
            phoneNumbers.append(cell.value)
        elif (count is 3):
            customFields.append(cell.value)
        elif (count is 4):
            messages.append(cell.value)
        elif (count is 5):
            approval.append(cell.value)

    count+=1

#Functions
def sendText(msg, phonenumber, clnt):
    message = clnt.messages.create(to="+" + str(phonenumber), from_="+14088989675", body=msg) #this one line sends the message
    return

def createMessage(message):
    header = ""
    footer = "\n-Curry on Wheels"
    response = message + footer
    return response

def isValid(index):
    #this section is modular and cam be customized to whatever counts as "valid"
    if (phoneNumbers[index] != None and messages[index] != None and approval[index] == "yes"):
        return True
    else:
        return False

#Counts valid messages
for i in range(1,len(phoneNumbers)):
    if(isValid(i)):
        countOfMessages+=1

#Asks for confirmation and takes user input
confirmation = raw_input(str(countOfMessages) + " messages will be sent. Continue? (y/n):")

log.info(str(countOfMessages)+" messages to be sent.")
log.info('User Entered ' + confirmation)

#Sends texts upon confirmation
if(confirmation=='y'):
    print "Sending messages...\n"
    for i in range(1, len(phoneNumbers)):
        if (isValid(i)): #checks for validity of message
            sendText(createMessage(messages[i]),phoneNumbers[i],client)
            log.info("Index: " + str(i) + " , Sent!")
        else:
           log.warning("Index: " + str(i) + " Not sent!")
else:
    print "Program Completed."