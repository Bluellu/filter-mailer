import tkinter as tk
from tkinter import messagebox
import smtplib as smtp
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from time import sleep
import io

def connect_server(user, pw, server_addr, port):
    #Connect to host
    server = None
    try:
        if port == '587':
            server = smtp.SMTP(server_addr, int(port))
            #server.set_debuglevel(1)
            server.starttls() #start encrypted session
            server.ehlo()
            print("connected to server")

        elif port == '465':
            server = smtp.SMTP_SSL(server_addr, int(port))
            #server.set_debuglevel(1)
            server.ehlo()
            print("connected to server") 
        else:
            messagebox.showinfo("Warning", "Port not supported. Please try 587 or 465.")
            return None
           
    except smtp.socket.gaierror:
        messagebox.showinfo("Error", "Could not contact host.")
        return None   

    #Log user in
    try:
        if server is not None:
            server.login(user, pw)
            print("Logged into server")
            
    except smtp.SMTPAuthenticationError:
        messagebox.showinfo("Error", "Could not authenticate user.")
        server.quit()
        return None

    #All good, return connected server
    finally:
        if server is not None:
            return server
        else:
            return None
    

def mail_all(subj, message, user, pw, server_addr, port, recipients):
    '''Send identical emails to every contact in the recipient list. '''

    #Build general email body
    msg = MIMEMultipart()
    msg['Subject'] = subj
    msg['From'] = user
    msg['To'] = ''
    msg.attach(MIMEText(message, 'plain'))
    
    server = connect_server(user, pw, server_addr, port)

    #Send emails
    try:
        if server is not None:
            if recipients:
                for email in recipients:
                    del msg['To']
                    msg['To'] = email
                    server.sendmail(user, email, msg.as_string())
                    print("Sent to " + email)
                   # sleep(20) #To only send 3 emails per minute
                    
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




        

    
