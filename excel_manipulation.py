# Excel manipulation functions
import tkinter as tk
from tkinter.filedialog import askopenfilename
from os.path import normpath, basename
from collections import OrderedDict
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

    # Remove duplicates 
    ordered_email_lst = list(OrderedDict.fromkeys(email_lst))
    print(ordered_email_lst, len(ordered_email_lst)) #TEST
    
    return ordered_email_lst

    #TODO: allow user to select correct email column if multiple available


def contains_term(email, terms):
    """ Returns a boolean indicating whether an email contains any of the given
        sub-terms. Helper for filter_emails.

    """
    contains = False
    for term in terms:
        if term.lower() in email.lower():
            contains = True
            break
    return contains

def filter_emails(email_lst, include_lst, exclude_lst):
    """ Separates an email list into approved and rejected email lists.

        Args:
            email_lst: List of all emails to be analyzed.
            include_lst: List of terms to be included. If empty, all items match.
            exclude_lst: List of terms to be filtered out.
        Returns:
            approved_ems: Emails that passed the filter.
            rejected_ems: Emails to be excluded.
        
    """
    approved_ems = []
    rejected_ems = []

    for email in email_lst:
        in_approve = (not include_lst) or contains_term(email, include_lst)
        in_exclude = contains_term(email, exclude_lst)

        if in_approve and (not in_exclude):
            approved_ems.append(email)
        else:
            rejected_ems.append(email)
    
    return approved_ems, rejected_ems




        
        
        
                    
        
        

    
