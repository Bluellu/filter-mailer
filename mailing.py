from os.path import normpath, basename
from time import sleep
import io
import datetime as dt
import tkinter as tk
from tkinter import messagebox
import smtplib as smtp
from email.message import EmailMessage

import pandas as pd


def connect_server(user, pw, server_addr, port):
    ''' Connect to a given server and log the given user in. '''
    
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
            messagebox.showinfo("Warning", "Port not supported. "
                                + "Please try 587 or 465.")
            return None
           
    except smtp.socket.gaierror:
        messagebox.showinfo("Error", "Could not contact host.")
        return None
    
    except smtp.SMTPConnectError:
        messagebox.showinfo("Error", "Connection failed.")
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
    
    except smtp.SMTPServerDisconnected:
        messagebox.showinfo("Warning", "Server disconnected (in connect_server).")
        return None

    #All good, return connected server
    finally:
        return server
        

def construct_email(subj, message, img_path, user):
    ''' Construct a base email with blank 'To' field for use by mail_all.
        It includes HTML and plaintext alternatives to the message, with
        an optional image attachment.
        
    '''    
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
            """
    html2 = """
                  </body>
                </html>
            """
    html_msg = """"""
    msg_w_br = message.replace('\n', '<br>').replace('\r', '')
    
    #Include image attachment   
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
        if message is not '': #with message
            html_msg = html1 +  """<p>"""+ msg_w_br +"""</p><br>
                                   <img src="cid:{img_cid}">""".format(
                                img_cid = img_cid) + html2

        else: #without message
            html_msg = html1 + """<img src="cid:{img_cid}">""".format(
                                img_cid = img_cid) + html2
            
        msg.add_alternative(html_msg, subtype = 'html')

        #Attach image
        msg.get_payload()[1].add_related(img_bin, 'image', img_type,
                                     cid = img_cid)
        
    #No-image version of HTML  
    else:
        if message is not '':
            html_msg = html1 + """<p>"""+ msg_w_br +"""</p>""" + html2 
            msg.add_alternative(html_msg, subtype = 'html')

    return msg
        

def mail_all(subj, message, img_path, user, pw, server_addr, port, recipients):
    ''' Send identical emails to every contact in the recipient list. '''
    print(port)
    msg = construct_email(subj, message, img_path, user)
   
    #Start server    
    server = connect_server(user, pw, server_addr, port)

    #Create backup excel file to keep track of sending status
    dt_now = dt.datetime.now().strftime("%d-%m-%y__%H-%M")
    backup_fn = 'backup\\backup_'+ dt_now +'.xlsx' #timestamped filename
    
    xl_writer = pd.ExcelWriter(backup_fn)
    data = {'Sent_to': recipients, 'Unsent_to': '' * len(recipients)}
    df = pd.DataFrame(data, columns = ['Sent_to', 'Unsent_to'])

    df.to_excel(xl_writer, 'sheet1', engine = 'openpyxl')
    xl_writer.save()

    #Send email to all recipients  
    if server is not None:
        if recipients:
            for email in recipients:
                try:
                    del msg['To']
                    msg['To'] = email
                    server.send_message(msg)
                    print("Sent to " + email)
                    #sleep(20) #Only send 3 emails per minute
                        
                except smtp.SMTPRecipientsRefused as e:
                    print(e)
                    messagebox.showinfo("Error", "The following recipient could not "
                                        + "receive your message: "
                                        + ''.join(list(e.recipients.keys())))
                    
                except smtp.SMTPDataError as e:
                    print(e)
                    messagebox.showinfo("Error", "Unexpected Error code.")

                except smtp.SMTPServerDisconnected:
                    messagebox.showinfo("Warning", "Server disconnected (in mail_all).")                    
            
            status = server.noop()
            if status[0] == 250:
                server.quit()
                messagebox.showinfo("Result", "All done.")





        

    
