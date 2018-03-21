# Excel manipulation functions
from tkinter import *
from tkinter.filedialog import askopenfilename
from os.path import normpath, basename
import xlrd
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
    xl_workbook = xlrd.open_workbook(app.filepath)

    # Assume input workbook has a single sheet #(TODO: Allow for multi-sheet files)    
    xl_sheet = xl_workbook.sheet_by_index(0)
    print ('Sheet: %s' % xl_sheet.name) #TEST

    #TODO: extract email list from column (use regular expression to check column name)
    #TODO: Add support for .xls type

    
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

    
