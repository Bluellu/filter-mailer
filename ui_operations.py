import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import ttk
from os.path import normpath, basename
import tkinter.scrolledtext as tkscrolled
import excel_manipulation as em
import mailing as ml

'''Functions that modify and create UI elements. '''

def center_window(window):
    ''' Center window within the screen. '''
    
    screen_h = window.winfo_screenheight()
    screen_w = window.winfo_screenwidth()
    win_h = window.winfo_height()
    win_w = window.winfo_width()

    x = (screen_w / 2) - (win_w / 2)
    y = (screen_h / 2) - (win_h / 2)

    window.geometry('%dx%d+%d+%d' % (win_w, win_h, x, y))

    
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
        

def attach_image(img_path):
    ''' Set attached-image label (helper for email_creation). '''    
    img_path.config(text = askopenfilename(initialdir = "C:/",
                                title = "import file",
                                filetypes =
                                [('Image', ('*jpeg', '*.jpg')),
                                ('Image', '*.png',)]))
def email_creation(app):
    ''' Construct email-creation window. '''
    
    ec = tk.Toplevel(app.frame, bd = 15)
    ec.wm_title("Email Creation")
    ec.geometry("450x400")    
    ec.update()
    center_window(ec)    
    ec.grab_set()

    recipients = em.get_sorted_emails(app)[0] #Approved emails

    #Subject
    subj_lbl = tk.Label(ec, text = "Subject")
    subj_box = tk.Entry(ec)

    #Message
    msg_lbl = tk.Label(ec, text = "Message")
    msg_box = tkscrolled.ScrolledText(ec, height = 10)

    #Selected image attachment path
    img_path = tk.Label(ec)

    #Attachment button
    img_bttn = tk.Button(ec, text = "Attach Image", width = 0,
                        fg = "white",
                        bg = "navy",
                        command = (lambda: attach_image(img_path)))
    
    #Login/Send button
    login_bttn = tk.Button(ec, text = "Log in and send emails", width = 70,
                        fg = "white",
                        bg = "navy",
                        command = (lambda: login_and_send(ec, subj_box.get(),
                                                              msg_box.get("1.0", 'end-1c'),
                                                              img_path.cget("text"),
                                                              recipients)))

    #Sender box
    recip_lbl = tk.Label(ec, text = "Recipients", bg = "light blue")
    recip_box = create_list_wgt(ec, recipients, 20, 15)
    
    #Add items to grid
    recip_lbl.grid(column = 2, row = 0, sticky = "nsew")
    recip_box.grid(column = 2, row = 1, rowspan = 3, sticky = "nsew")

    subj_lbl.grid(column = 0, row = 0, columnspan = 2, padx = 10, sticky = "w")
    subj_box.grid(column = 0, row = 1, columnspan = 2, padx = (10, 25), sticky = "ew")

    msg_lbl.grid(column = 0, row = 2, columnspan = 2, padx = 10, sticky = "w")
    msg_box.grid(column = 0, row = 3, columnspan = 2, padx = 10, sticky = "nsew")

    img_path.grid(column = 1, row = 4, columnspan = 3, sticky = "w")
    img_bttn.grid(column = 0, row = 4, padx = 10, pady = 10, sticky = "w")
    
    login_bttn.grid(column = 0, row = 5, columnspan = 3, padx = 50, pady = 10, sticky = "ns")

    ec.columnconfigure(0, weight = 0)
    ec.columnconfigure(1, weight = 1)
    ec.rowconfigure(3, weight = 1)
    

def login_and_send(ec, subj, msg, img_path, recipients):
    ''' Construct login window. Child of email-creation window. '''

    ls = tk.Toplevel(ec, bd = 15)
    ls.wm_title("Log in")
    ls.geometry("300x200")
    ls.update()
    center_window(ls)    
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
                        command = (lambda: ml.mail_all(subj,
                                                        msg,                                                       
                                                        img_path,
                                                        email_box.get(),
                                                        pw_box.get(),
                                                        server_box.get(),
                                                        port_box.get(),
                                                        recipients)))

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


    



    
    


    
