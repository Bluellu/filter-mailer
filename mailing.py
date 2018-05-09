import tkinter as tk
from tkinter import messagebox
import smtplib as smtp
from email.mime.text import MIMEText
from time import sleep
import io

def mass_mail(subj, message, user, pw, server_addr, port, recipients):
    '''Send identical emails to every contact in the recipient list. '''
    #Connect to host
    try:
        server = None
        if port == '587':
            server = smtp.SMTP(server_addr, int(port))
            server.set_debuglevel(1)
            server.starttls() #start encrypted session
            server.ehlo()
            print("connected to server")        

            print("Starting encrypted session")
        elif port == '465':
            server = smtp.SMTP_SSL(server_addr, int(port))
            server.set_debuglevel(1)
            server.ehlo()
            print("connected to server") 
        else:
            messagebox.showinfo("Warning", "Port not supported. Please try 587 or 465.")
            return False
           
    except smtp.socket.gaierror:
        messagebox.showinfo("Error", "Could not contact host.")
        return False    

    #Log user in
    try:
        if server is not None:
            server.login(user, pw)
            print("Logged into server.")
    except smtp.SMTPAuthenticationError:
        messagebox.showinfo("Error", "Could not authenticate user.")
        server.quit()
        return False

    #Send emails
    try:
        if server is not None:
            ems = [''] #TEST
            for email in ems:
            #for email in recipients:
                server.sendmail(user, email, 'Hello, this is a test.') #TEST
    except smtp.SMTPRecipientsRefused as e:
        print(e)
        messagebox.showinfo("Error", "The following recipient could not receive your message: "
                            + ''.join(list(e.recipients.keys())))
    except smtp.SMTPDataError as e:
        print(e)
        messagebox.showinfo("Error", "Unexpected Error code.")    

    finally:
        if server is not None:
            server.quit()
            messagebox.showinfo("Result", "All done.")
            print("Done!")

    #Build email body  
##        msg = MIMEText(message)
##        msg['Subject'] = subj
##        msg['From'] = user
##        msg['To'] = '###' #TEST

    #Send message to each recipient 
##        for email in recipients:
##        msg['To'] = email
##            send_message(msg, user, to_addr)
##            print("Sent to " + email)
##            sleep(20) #To only send 3 emails per minute



        

    
