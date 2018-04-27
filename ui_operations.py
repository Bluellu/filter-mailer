import tkinter as tk
from tkinter.filedialog import askopenfilename
from os.path import normpath, basename
import excel_manipulation as em
import tkinter.scrolledtext as tkscrolled

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

    #Image attachment
    img_bttn = tk.Button(ec, text = "Attach Image", width = 0,
                        fg = "white",
                        bg = "navy",
                        #command = (lambda: attach_image(ec))
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

    ec.columnconfigure(0, weight = 1)
    ec.rowconfigure(3, weight = 1)




    
    


    
