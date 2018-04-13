import tkinter as tk
from tkinter.filedialog import askopenfilename
from os.path import normpath, basename
import excel_manipulation as em

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

def filter_preview(app):
    """ Display filtered version of the email list in new window. """
    
    #Create pop-up window
    preview = tk.Toplevel(app.frame)
    preview.wm_title("Filter Preview")

    if (app.filepath):
        filtered_emails = em.get_sorted_emails(app)

        accepted_lst = tk.Text(preview, width = 15, height = 20)
        for email in filtered_emails[0]:
            accepted_lst.insert(tk.END, email + '\n')
        accepted_lst.grid(column = 0, row = 1)

        rejected_lst = tk.Text(preview, width = 15, height = 20)
        for email in filtered_emails[1]:
            rejected_lst.insert(tk.END, email + '\n')
        rejected_lst.grid(column = 1, row = 1)         
        
##        print('APPROVED/////////////////////')
##        print(filtered_emails[0])
##        print('REJECTED////////////////////')
##        print(filtered_emails[1])
##        print('END/////////////////////////////////////////////////////////////////')
    
    else:
        warning = tk.Label(preview, text = "No File Selected")
        warning.grid()
    


    
