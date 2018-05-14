import tkinter as tk
from tkinter import messagebox
import smtplib as smtp
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from os.path import normpath, basename
from time import sleep
import io

from email.message import EmailMessage
from email.utils import make_msgid
import mimetypes

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
    

def mail_all(subj, message, user, pw, server_addr, port, img_path, recipients):
    '''Send identical emails to every contact in the recipient list. '''

    #Base email body
    msg = EmailMessage()
    msg['Subject'] = subj
    msg['From'] = user
    msg['To'] = ''    
    msg.set_content(message) #Add plaintext version of the message

    #Base HTML body
    html1 = """ 
                <html>
                  <head></head>
                  <body>
                    <p>"""+ message +"""</p><br>"""
    html2 = """
                  </body>
                </html>
            """
    html_msg = """"""
    
    #Image attachment   
    if img_path.strip() is not '':
        
        #Get file path
        fp = open(img_path, 'rb')
        img_bin = fp.read()
        fp.close()
        
        file_name = basename(normpath(img_path))
        name_split = file_name.rsplit('.', 1)
        img_type = name_split[len(name_split)-1] #file type
        img_cid = name_split[len(name_split)-2] #file name (no extension)

        #Final HTML message with image ID
        html_msg = html1 + """<img src="cid:{img_cid}">""".format(
                            img_cid = img_cid) + html2
        msg.add_alternative(html_msg, subtype = 'html')

        #Attach image
        msg.get_payload()[1].add_related(img_bin, 'image', img_type,
                                     cid = img_cid)
        
    else:
        html_msg = html1 + html2 #no-image version of HTML
        msg.add_alternative(html_msg, subtype = 'html')
        #msg.make_mixed() #convert to multipart/mixed


    #Start server and send email to all recipients     
    server = connect_server(user, pw, server_addr, port)
    
    try:
        if server is not None:
            if recipients:
                for email in recipients:
                    del msg['To']
                    msg['To'] = email
                    server.send_message(msg)
                    print("Sent to " + email)
                    sleep(20) #Only send 3 emails per minute
                    
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




        

    
