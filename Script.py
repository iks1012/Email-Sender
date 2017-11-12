import smtplib
import os, sys
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import User
from email import encoders

# Email You are using

user = User()

sendingEmail = "isethi@cs.stonybrook.edu"
password = User.getPass()

hwNum = 5
sendingHW = True

TAFolderEXT = "G:\\Work Related\\Education\\College\\Sophomore Year\\01 - Fall\\CSE 214 TA\\"
studentData = "Students.txt"
HWFolder = "HW" + str(hwNum) + " - R06\\"

hwPath = TAFolderEXT + HWFolder
fileNames = os.listdir(hwPath)

# Login to Gmail
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sendingEmail, password)

# Load up sending emails
studentsTXT = open(TAFolderEXT + studentData)
stringFormOfFile = studentsTXT.read()
students = stringFormOfFile.split('\n')

fileOffset = 0

# Send ith Attachment to ith email with the message
for i in range(len(students) - 1):
    x = students[i].index("-") + 1
    tempEmail = students[i][x:].strip()
    tempFileName = fileNames[i - fileOffset]

    if (tempEmail[:tempEmail.index(".")] in tempFileName.lower()) or (
        tempEmail[tempEmail.index(".") + 1:tempEmail.index("@")] in tempFileName.lower()):
        msg = "This Script actually works really well :O \n" \
                "Attached is the homework grade sheets for homework " + str(hwNum) + ".\n" \
                "If there any discrepancies about the grade please contact the grading TA.\n" \
                "If there are any immediate issues please let me know.\n"
        if sendingHW:
            msg2 = MIMEMultipart()
            msg2['From'] = sendingEmail
            msg2['To'] = tempEmail
            msg2['Subject'] = "Homework " + str(hwNum) + " Grade Sheet"
            msg2.attach(MIMEText(msg, 'plain'))

            attachment = open(hwPath + tempFileName, "rb")

            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', "attachment; filename= %s" % tempFileName)
            msg2.attach(part)
            msg = msg2.as_string()

        server.sendmail(sendingEmail, tempEmail, msg)

        print("(" + sendingEmail + " -> " + tempEmail + "): " + tempFileName)
    else:
        fileOffset += 1

server.quit()
