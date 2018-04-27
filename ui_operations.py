import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import ttk
from os.path import normpath, basename
import tkinter.scrolledtext as tkscrolled
import excel_manipulation as em
#import mailing as ml

'''Functions that modify and create UI elements. '''

def get_filepath(app, label):
    """ Set the app's filepath to a chosen file and
        update the GUI label to display the chosen file's name.

        Args:
            app: App instance containing the filepath argument to be set.
            label: GUI label to be updated to display the new file's name.            
        
    """
    app.filepath = askopenfilename(initialdir = "C:/",
                                title = "import file",
                                filetypes =
                                [('Excel', ('*.xls', '*.xlsx')),
                                ('CSV', '*.csv',)])
    
    # Update displayed file name
    if (app.filepath):
        label.config(text = basename(normpath(app.filepath)))
        #print(app.filepath)


def create_list_wgt(app, txt_list, x, y):
    '''Return a text widget that displays the given list, with size x by y.'''
    txt_widget = tkscrolled.ScrolledText(app, width = x, height = y)    
    for item in txt_list:
        txt_widget.insert(tk.END, item + '\n')
        
    return txt_widget
    
def filter_preview(app):
    """ Display filtered version of the email list in new window. """
    
    #Create pop-up window
    prvw = tk.Toplevel(app.frame)
    prvw.wm_title("Filter Preview")

    if (app.filepath):
        filtered_emails = em.get_sorted_emails(app)

        accepted_lst = create_list_wgt(prvw, filtered_emails[0], 20, 20)
        accepted_lst.grid(column = 0, row = 1)

        rejected_lst = create_list_wgt(prvw, filtered_emails[1], 20, 20)
        rejected_lst.grid(column = 1, row = 1)         
    
    else:
        warning = tk.Label(prvw, text = "No File Selected")
        warning.grid()

def login_and_send(ec, subj_box, msg_box, recipients):
    ls = tk.Toplevel(ec, bd = 15)
    ls.wm_title("Log in")
    ls.geometry("300x200")
    
    ls.grab_set()

    #Server input
    server_lbl = tk.Label(ls, text = "Server: ")
    server_box = tk.Entry(ls)

    #Port input
    port_lbl = tk.Label(ls, text = "Port: ")
    port_box = tk.Entry(ls)

    #User email input
    email_lbl = tk.Label(ls, text = "Email: ")
    email_box = tk.Entry(ls)

    #Password input
    pw_lbl = tk.Label(ls, text = "Password: ")
    pw_box = tk.Entry(ls, show = "*")

    #Send button
    send_bttn = tk.Button(ls, text = "Send emails", width = 30,
                        fg = "white",
                        bg = "navy",
                        #command = (lambda: login_and_send(ec, subj_box, msg_box, recipients))
                        )

    sep = ttk.Separator(ls, orient = "horizontal")
    
    server_lbl.grid(column = 0, row = 0, padx = 10, pady = 5, sticky = "w")
    port_lbl.grid(column = 0, row = 1, padx = 10, pady = 5, sticky = "w")
    email_lbl.grid(column = 0, row = 3, padx = 10, pady = 5, sticky = "w")
    pw_lbl.grid(column = 0, row = 4, padx = 10, pady = 5, sticky = "w")

    sep.grid(column = 0, row = 2, columnspan = 2, sticky = "nsew")

    server_box.grid(column = 1, row = 0, padx = 10, pady = 5, sticky = "nsew")
    port_box.grid(column = 1, row = 1, padx = 10, pady = 5, sticky = "nsew")
    email_box.grid(column = 1, row = 3, padx = 10, pady = 5, sticky = "nsew")
    pw_box.grid(column = 1, row = 4, padx = 10, pady = 5, sticky = "nsew")

    send_bttn.grid(column = 0, row = 5, columnspan = 2, padx = 20, pady = 10, sticky = "ns")
    
    

    #ml.mass_send(#TODO user, pw, server_addr, port, recipients)

def email_creation(app):
    ec = tk.Toplevel(app.frame, bd = 15)
    ec.wm_title("Email Creation")
    ec.geometry("450x400")
    
    ec.grab_set()

    recipients = em.get_sorted_emails(app)[0] #Approved emails

    #Subject
    subj_lbl = tk.Label(ec, text = "Subject")
    subj_box = tk.Entry(ec)

    #Message
    msg_lbl = tk.Label(ec, text = "Message")
    msg_box = tkscrolled.ScrolledText(ec, height = 10)

    #Attachment button
    img_bttn = tk.Button(ec, text = "Attach Image", width = 0,
                        fg = "white",
                        bg = "navy",
                        #command = (lambda: attach_image(ec))
                        )
    #Login/Send button
    login_bttn = tk.Button(ec, text = "Log in and send emails", width = 70,
                        fg = "white",
                        bg = "navy",
                        command = (lambda: login_and_send(ec, subj_box, msg_box, recipients))
                        )

    #Sender box
    recip_lbl = tk.Label(ec, text = "Recipients", bg = "light blue")
    recip_box = create_list_wgt(ec, recipients, 20, 15)
    
    #Add items to grid
    recip_lbl.grid(column = 1, row = 0, sticky = "nsew")
    recip_box.grid(column = 1, row = 1, rowspan = 4, sticky = "nsew")

    subj_lbl.grid(column = 0, row = 0, padx = 10, sticky = "w")
    subj_box.grid(column = 0, row = 1, padx = (10, 25), sticky = "ew")

    msg_lbl.grid(column = 0, row = 2, padx = 10, sticky = "w")
    msg_box.grid(column = 0, row = 3, padx = 10, sticky = "nsew")

    img_bttn.grid(column = 0, row = 4, padx = 10, pady = 10, sticky = "w")
    login_bttn.grid(column = 0, row = 5, columnspan = 2, padx = 50, pady = 10, sticky = "ns")

    ec.columnconfigure(0, weight = 1)
    ec.rowconfigure(3, weight = 1)

    



    
    


    
