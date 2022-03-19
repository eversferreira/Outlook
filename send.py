import os
import win32com.client as win32


def sendmail():
    # Mail variables
    mail_to = 'ADDRESSEE'
    mail_subject = 'SUBJECT'
    mail_msg = f""" 
                <P>YOUR HTML MESSAGE</p>"""

    outlook = win32.Dispatch('Outlook.Application')
    mail = outlook.CreateItem(0)
    mail.To = mail_to
    mail.Subject = mail_subject
    mail.BodyFormat = 1
    mail.HTMLBody = mail_msg

    # Attachments
    # Important: Have sure you placed the files on project folder
    # Clone this line and change the file name if you need multiple attachments
    mail.Attachments.Add(os.path.join(os.getcwd(), 'FILENAME.EXTENSION'))

    # If you want to use the Outlook interface
    # mail.Display()

    mail.Send()

