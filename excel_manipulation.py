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
        

    
def filter_emails(email_lst, filter_list, exclude):
    """ Separates an email list into two lists of approved and rejected emails.

        Args:
            email_lst: List of all emails to be analyzed.
            filter_lst: List of terms to be filtered in or out of approval list.
            exclude: Bolean to indicate whether the filter list approves or excludes an email string.
        Returns:
            approved_lst: Emails that passed the filter.
            rejected_lst: Emails to be excluded.
        
    """

    
