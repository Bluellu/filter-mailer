# Excel manipulation functions
import tkinter as tk
from tkinter.filedialog import askopenfilename
from os.path import normpath, basename
#import xlrd
import pandas as pd
import re


def get_filepath(app, label):
    """ Sets the app's filepath to a chosen file and
        updates the GUI label to display the chosen file's name.

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
        print(app.filepath)

        extract_email_lst(app) #TEST
        
        
def extract_email_lst(app):
    """ Extracts and returns a list of emails from an excel file column.

        Args:
            app: The app instance containing the desired excel filepath.
        Returns:
            A single list of email strings.
            
    """
    df = pd.read_excel(app.filepath)
    email_lst = []
    
    col_name = 'E-Mail'

    if col_name in df.columns.values:
        email_lst = df['E-Mail'].values #TODO: Allow different spellings with regex
        print(email_lst, len(email_lst)) #TEST

    return email_lst
    #TODO: allow user to select correct email column if multiple available


def filter_emails(filter_lst, email_lst, exclude):
    """ Separates an email list into approved and rejected email lists.

        Args:
            email_lst: List of all emails to be analyzed.
            filter_lst: List of terms to be filtered in or out of approval list.
            exclude: Bolean to indicate whether the filter list approves or excludes an email string.
        Returns:
            approved_ems: Emails that passed the filter.
            rejected_ems: Emails to be excluded.
        
    """
    approved_ems = []
    rejected_ems = []

    for email in email_lst:
        #Check if email matches any filter words
        in_filter = False
        for filter_i in filter_lst:
            in_filter = filter_i.lower() in email.lower()
            print(filter_i.lower(), email.lower(), in_filter)
            
        if in_filter == exclude:
            rejected_ems.append(email)
        elif in_filter != exclude:
            approved_ems.append(email)
    
    return approved_ems, rejected_ems



        
        
        
                    
        
        

    
