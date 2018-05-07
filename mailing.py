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
        server = smtp.SMTP(server_addr, port)
        server.ehlo()
        print("connected to server")        
        server.starttls() #start encrypted session
        server.ehlo()
        print("Starting encrypted session")
    except smtp.socket.gaierror:
        messagebox.showinfo("Error", "Could not contact host.")
        return False    

    #Log user in
    try:
        server.login(user, pw)
        print("Logged into server.")
    except smtp.SMTPAuthenticationError:
        messagebox.showinfo("Error", "Could not authenticate user.")
        server.quit()
        return False

    #Send emails
    try:
        # ems = [] #TEST
        #for email in ems:
        for email in recipients:
            server.sendmail(user, email, 'TESTING CODE') #TEST
    #TODO
    except Exception as e:
        print(e)

    finally:
        if server is not None:
            server.quit()    

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



        

    
