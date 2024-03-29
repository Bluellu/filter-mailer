'''Functions that create and modify UI elements. '''
from os.path import normpath, basename
import tkinter as tk
from tkinter.filedialog import askopenfilename
from tkinter import ttk
import tkinter.scrolledtext as tkscrolled
from tkinter import messagebox
from time import sleep

import excel_manipulation as em
import mailing as ml

class StatusHandler:
    ''' Creates and manages a status-display window to be used
    within mail_all. '''
    
    def __init__(self, parent, num_iters):
        self.root = parent
        self.iters = num_iters
        self.curr_iter = 0
        self.cancellation_status = False
        
        self.frame = tk.Toplevel(self.root, bd = 15) 
        self.frame.wm_title("Sending Emails: Status")
        self.frame.geometry("300x200")

        self.frame.update()
        center_window(self.frame) 
   
        self.frame.grab_set()
        
        self.text = subj_lbl = tk.Label(self.frame, text = "Beginning....")

        # Cancel button (to stop program)
        cancel_bttn = tk.Button(self.frame, text = "Cancel", width = 10,
                        fg = "white",
                        bg = "navy",
                        command = self.cancel)

        self.text.grid(column = 0, row = 0, pady = 35)
        cancel_bttn.grid(column = 0, row = 1)
        self.frame.columnconfigure(0, minsize = 250)

        self.frame.update()
        center_window(self.frame)

    def cancel(self):
        self.cancellation_status = True

    def cancelled(self):
        return self.cancellation_status

    def self_destruct(self):
        self.root.grab_set()
        self.frame.destroy()

    def updateMessage(self, recipient):
        self.curr_iter += 1
        self.text.config(text = "Sending email " + str(self.curr_iter) + " of "
                         + str(self.iters) + " to:\n" + recipient)
        self.frame.update()

        # Last update. Return control to parent window and self-destruct.
        if (self.curr_iter == self.iters):
            sleep(5)
            self.root.grab_set()
            self.frame.destroy()
            
        
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
        app(App): Main application instance containing the filepath argument.
        label(tk.Label): GUI label to be updated to display the new file's name.            
        
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
    ''' Return a text widget that displays the given list, with size x by y.'''
    txt_widget = tkscrolled.ScrolledText(app, width = x, height = y)    
    for item in txt_list:
        txt_widget.insert(tk.END, item + '\n')

    default_colour = app.cget('bg')
    txt_widget.configure(bg = default_colour, state = tk.DISABLED)
        
    return txt_widget
    
def filter_preview(app):
    """ Display filtered version of the email list in new window. """
    
    # Create pop-up window
    prvw = tk.Toplevel(app.frame)
    prvw.wm_title("Filter Preview")

    if (app.filepath):
        filtered_emails = em.get_sorted_emails(app)

        accepted_lbl = tk.Label(prvw, text = "Accepted", fg = "navy", bg = "light blue")
        rejected_lbl = tk.Label(prvw, text = "Rejected", fg = "navy", bg = "light blue")

        accepted_lst = create_list_wgt(prvw, filtered_emails[0], 20, 20)
        rejected_lst = create_list_wgt(prvw, filtered_emails[1], 20, 20)
        
        accepted_lbl.grid(column = 0, row = 1, sticky='ew')
        rejected_lbl.grid(column = 1, row = 1, sticky='ew')
        accepted_lst.grid(column = 0, row = 2)
        rejected_lst.grid(column = 1, row = 2)        
    
    else:
        warning = tk.Label(prvw, text = "No File Selected")
        warning.grid()        


def email_creation(app):
    ''' Construct email-creation window. '''
    
    ec = tk.Toplevel(app.frame, bd = 15)
    ec.wm_title("Email Creation")
    ec.geometry("450x400")    
    ec.update()
    center_window(ec)    
    ec.grab_set()

    recipients = em.get_sorted_emails(app)[0] # Approved emails

    # Subject
    subj_lbl = tk.Label(ec, text = "Subject")
    subj_box = tk.Entry(ec)

    # Message
    msg_lbl = tk.Label(ec, text = "Message")
    msg_box = tkscrolled.ScrolledText(ec, height = 10)

    # Selected image attachment path
    img_path = tk.Label(ec)

    # Image attachment button setup
    def attach_image(img_path):
        ''' Set attached-image label (helper for email_creation). '''    
        img_path.config(text = askopenfilename(initialdir = "C:/",
                                    title = "import file",
                                    filetypes =
                                    [('Image', ('*jpeg', '*.jpg')),
                                    ('Image', '*.png',)]))
    
    img_bttn = tk.Button(ec, text = "Attach Image", width = 0,
                        fg = "white",
                        bg = "navy",
                        command = (lambda: attach_image(img_path)))
    
    # Login/Send button
    login_bttn = tk.Button(ec, text = "Log in and send emails", width = 70,
                        fg = "white",
                        bg = "navy",
                        command = (lambda: login_and_send(app, ec, subj_box.get(),
                                                              msg_box.get("1.0", 'end-1c'),
                                                              img_path.cget("text"),
                                                              recipients)))
    # Sender box
    recip_lbl = tk.Label(ec, text = "Recipients", bg = "light blue")
    recip_box = create_list_wgt(ec, recipients, 20, 15)
    #default_colour = ec.cget('bg')
    #recip_box.configure(bg = default_colour, state = tk.DISABLED)
    
    # Add items to grid
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
    

def login_and_send(app, ec, subj, msg, img_path, recipients):
    ''' Construct login window. Child of email-creation window. '''

    if (msg is '') and (img_path is ''):
        messagebox.showinfo("Invalid input", "Please write a message and/or "
                            + "attach an image.")
        
    else: # create login box
        ec.withdraw() #Hide parent window
        ls = tk.Toplevel(ec, bd = 15)
        ls.wm_title("Log in")
        ls.geometry("300x200")
        ls.update()
        center_window(ls)    
        ls.grab_set()

        def ls_quit():
            ec.deiconify()
            ls.destroy()            

        ls.protocol("WM_DELETE_WINDOW", ls_quit)

        # Server input
        server_lbl = tk.Label(ls, text = "Server: ")
        server_box = tk.Entry(ls)

        # Port input
        port_lbl = tk.Label(ls, text = "Port: ")
        value = '587'
        port_box = ttk.Combobox(ls, textvariable = value, state = 'readonly')
        port_box['values'] = ('587', '465')
        port_box.current(0)

        # User email input
        email_lbl = tk.Label(ls, text = "Email: ")
        email_box = tk.Entry(ls)

        # Password input
        pw_lbl = tk.Label(ls, text = "Password: ")
        pw_box = tk.Entry(ls, show = "*")

        # Send button setup
        def send_handler():
            ls.withdraw() 
            stat_handler = StatusHandler(app.frame, len(recipients))
            success = ml.mail_all(subj, msg, img_path, email_box.get(),
                                                            pw_box.get(),
                                                            server_box.get(),
                                                            port_box.get(),
                                                            recipients,
                                                            stat_handler)
            # Return to main app window
            ls.destroy()
            ec.destroy()
            
        send_bttn = tk.Button(ls, text = "Send emails", width = 30,
                            fg = "white",
                            bg = "navy",
                            command = send_handler)

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
    



    
    


    
