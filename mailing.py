from os.path import normpath, basename
from time import sleep
import io, os
import tkinter as tk
from tkinter import messagebox
import smtplib as smtp
from email.message import EmailMessage
from email.mime.text import MIMEText

import excel_manipulation as em


def connect_server(user, pw, server_addr, port):
    ''' Connect to a given server and log the given user in.

    Args:
        user(str): The user's email address.
        pw(str): The user's password.
        server_addr(str): SMTP server to connect to.
        port(str): Port number for SMTP connection.
    '''
    
    # Connect to host
    server = None
    try:
        if port == '587':
            server = smtp.SMTP(server_addr, int(port))
            #server.set_debuglevel(1)
            server.starttls() # start encrypted session
            server.ehlo()

        elif port == '465':
            server = smtp.SMTP_SSL(server_addr, int(port))
            #server.set_debuglevel(1)
            server.ehlo()

        else:
            messagebox.showinfo("Warning", "Port not supported. "
                                + "Please try 587 or 465.")
            return None
           
    except smtp.socket.gaierror:
        messagebox.showinfo("Error", "Could not contact host.")
        server = None
    
    except smtp.SMTPConnectError:
        messagebox.showinfo("Error", "Connection failed.")
        server = None

    except TimeoutError:
        messagebox.showinfo("Timeout Error", "No response.")        

    # Log user in
    try:
        if server is not None:
            server.login(user, pw)
            print("Logged into server")
            
    except smtp.SMTPAuthenticationError:
        messagebox.showinfo("Error", "Could not authenticate user.")
        server.quit()
        server = None
    
    except smtp.SMTPServerDisconnected:
        messagebox.showinfo("Warning", "Server disconnected during login.")
        server = None

    # All good, return connected server
    finally:
        return server
        

def construct_email(subj, message, img_path, user):
    ''' Construct a base email with blank 'To' field for use by mail_all.
    It includes HTML and plaintext alternatives to the message, with
    an optional image attachment.        
    '''
    
    # Base email body
    msg = EmailMessage()
    msg['Subject'] = subj
    msg['From'] = user
    msg['To'] = ''    
    msg.set_content(message) # Add plaintext version of the message

    # Base HTML body
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
    
    # Include image attachment   
    if img_path.strip() is not '':
        
        # Get file path
        fp = open(img_path, 'rb')
        img_bin = fp.read()
        fp.close()
        
        file_name = basename(normpath(img_path))
        name_split = file_name.rsplit('.', 1)
        img_type = name_split[len(name_split)-1] # file type
        img_cid = name_split[len(name_split)-2] # file name (no extension)

        # Final HTML message with image ID
        if message is not '': # with message
            html_msg = html1 +  """<p>"""+ msg_w_br +"""</p><br>
                                   <img src="cid:{img_cid}">""".format(
                                img_cid = img_cid) + html2

        else: # without message
            html_msg = html1 + """<img src="cid:{img_cid}">""".format(
                                img_cid = img_cid) + html2
            
        msg.add_alternative(html_msg, subtype = 'html')

        # Attach image
        msg.get_payload()[1].add_related(img_bin, 'image', img_type,
                                     cid = img_cid)
        
    # No-image version of HTML  
    else:
        if message is not '':
            html_msg = html1 + """<p>"""+ msg_w_br +"""</p>""" + html2 
            msg.add_alternative(html_msg, subtype = 'html')

    return msg


def mail_final_report(user, server, msg, success_lst, fail_lst):
    ''' Emails a copy of the mass-mailed message to the sender (user) including
    a report of which recipients have received it or not (A success list, and a
    failure list as text attachments).

    Args:
        user: Sender's email.
        server: A connected SMTP server.
        msg: An EmailMessage object. 
        success_lst: A list of recipients towards which a successful attempt
            at sending msg has been made.
        fail_lst: List of recipients that could not the message.
    '''

    # TODO: include success/failure lists in msg
    # (either as text or an attachment).
    
    if server:
        del msg['To']
        msg['To'] = user
        
        temp_subj = msg['Subject'] + ' (sender copy)'
        del msg['Subject']
        msg['Subject'] = temp_subj
        
        success_str = ''
        for i in success_lst:
            success_str = success_str + i + '\n'

        fail_str = ''
        for i in fail_lst:
            fail_str = fail_str + i + '\n'        

        # Attach files to message
        att1 = MIMEText(success_str) 
        att1.add_header('Content-Disposition', 'attachment', filename='successes.txt')
        #msg.make_mixed() 
        msg.attach(att1)

        att2 = MIMEText(fail_str) 
        att2.add_header('Content-Disposition', 'attachment', filename='failures.txt')
        msg.attach(att2) 

        # Send message
        try:
            server.send_message(msg)
                        
        except smtp.SMTPRecipientsRefused as e:
            print(e)
            
        except smtp.SMTPDataError as e:
            print(e)

        except smtp.SMTPServerDisconnected:
            messagebox.showinfo("Warning", "Server disconnected.")

        # Delete temporary files         
        try:
            os.remove('success.txt')
            os.remove('failure.txt')           
        except OSError:
            pass
    

def mail_all(subj, message, img_path, user, pw, server_addr, port, recipients,
             stat_handler):
    ''' Send identical emails to every contact in the recipient list.

    Returns:
        bool: Success status.
    '''
    
    server = connect_server(user, pw, server_addr, port)    
    msg = construct_email(subj, message, img_path, user)

    backup_f = em.BackupFile(recipients) # excel file for send-status backup

    # Send email to all recipients
    if server is None:
        return False
    else:
        if recipients:
            success_lst = []
            failure_lst = []
            curr_i = 0
            for email in recipients:

                # Cancel button has been pressed. End of execution.
                if stat_handler.cancelled():
                    print("CANCELLED")

                    curr_index = recipients.index(email)
                    recipients_rest = recipients[curr_index:len(recipients)]
                    failure_lst = failure_lst + recipients_rest
                    stat_handler.self_destruct()
                    print(success_lst, failure_lst)
                    #mail_final_report(user, server, msg, success_lst, failure_lst)
                    break

                else:
                    stat_handler.updateMessage(email)
                    print("Sending to: " + email)
                    try:
                        del msg['To']
                        msg['To'] = email
                        server.send_message(msg)

                        # Update backup
                        backup_f.update(email, curr_i)
                        success_lst.append(email)
                        sleep(5) # Only send 3 emails per minute
                            
                    except smtp.SMTPRecipientsRefused as e:
                        print(e)
                        curr_i += 1 # Skip this index from unsent column
                        failure_lst.append(email)
                        
                    except smtp.SMTPDataError as e:
                        print(e)
                        curr_i += 1 # Same as above
                        failure_lst.append(email)

                    except smtp.SMTPServerDisconnected:
                        messagebox.showinfo("Warning", "Server disconnected.")                    

##            # Display warning about unsent emails
##            if failure_lst: 
##                failures = ', '.join(failure_lst) + '.'
##                    
##                messagebox.showinfo("Warning", "The following recipients "
##                                    + "could no receive your message:\n "
##                                    + failures) # TODO: Make this a scrollable widget

            mail_final_report(user, server, msg, success_lst, failure_lst)
            
            status = server.noop()
            if status[0] == 250:
                server.quit()
                messagebox.showinfo("Result", "All done.")
            return True





        

    
