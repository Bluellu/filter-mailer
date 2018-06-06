# Excel manipulation functions
import datetime as dt
from collections import OrderedDict
from collections import namedtuple

import pandas as pd

class BackupFile:
    ''' Encapsulates a dataframe and manipulation functions for an Excel backup
    file. This file stores recipients sorted by send-success status.
    '''
    
    def __init__(self, recipients):
        ''' Create new dataframe and xlsx file. All recipients are set in
        Unsent_to by default.

        Args:
            recipients: The list of emails to be stored in this backup file.
        '''
        
        # Dataframe to store backup information
        data = {'Unsent_to': recipients, 'Sent_to': ''}
        df = pd.DataFrame(data, columns = ['Unsent_to', 'Sent_to'])
        
        # Timestamped filename
        dt_now = dt.datetime.now().strftime("%d-%m-%y__%H-%M")
        backup_fn = 'backup\\backup_'+ dt_now +'.xlsx'

        # Writer to convert dataframe into an xlsx file
        xl_writer = pd.ExcelWriter(backup_fn)

        # Write initial file
        df.to_excel(xl_writer, 'sheet1', engine = 'openpyxl')
        xl_writer.save()

        # Init dataframe and excel writer for this file
        self.df = df
        self.writer = xl_writer
        self.curr_sent = 0 # Index of current (empty) Sent_to cell
        self.num_recip = len(recipients) # Total number of recipients

        
    def update(self, recipient, rm_index):
        ''' Add recipient to Sent_to column of backup spreadsheet,
        and remove it from Unsent_to column. '''

        # Add recipient to Sent_to column        
        self.df.Sent_to[self.curr_sent] = recipient
        self.curr_sent += 1 # Move index position to next free-cell
        
        # Remove recipient from Unsent_to column
        self.df.Unsent_to[rm_index:self.num_recip] = (
            self.df.Unsent_to[rm_index:self.num_recip].shift(-1).fillna(''))

##        with pd.option_context('display.max_rows', None, 'display.max_columns', 3):
##            print(self.df) #TEST
            
        # Ovewrite backup file
        self.df.to_excel(self.writer, 'sheet1', engine = 'openpyxl')
        self.writer.save()

        
def extract_email_lst(app):
    """ Extract and return a list of emails from an excel file column.

    Args:
        app(App): The app instance containing the desired excel filepath.
    Returns:
        A list of emails.
            
    """
    ordered_email_lst = []
    if (app.filepath):
        col_name = 'E-Mail'
        
        df = pd.read_excel(app.filepath)
     
        if col_name in df.columns.values:
            email_lst = df[df[col_name].notnull()][col_name].values #TODO: Allow different spellings with regex

            # Remove duplicates
            ordered_email_lst = list(OrderedDict.fromkeys(email_lst))
        
    return ordered_email_lst

    # TODO: allow user to select correct email column if multiple available


def contains_term(email, terms):
    """ Return a boolean indicating whether an email contains any of the given
        sub-terms. Helper for filter_emails.

    """
    contains = False
    if isinstance(email,(str,)):
        for term in terms:
            if isinstance(term,(str,)):
                if term.lower() in email.lower():
                    contains = True
                    break
    return contains

def filter_emails(email_lst, include_lst, exclude_lst):
    """ Separate an email list into approved and rejected email lists.

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


def process_list(string):
    """ Create list from string with items separated by commas. """
    # Remove all spaces
    no_spaces = "".join(string.split())

    # Split at commas
    split_lst = no_spaces.split(",")

    # Remove empty strings and return
    return list(filter(None, split_lst))
    
def get_sorted_emails(app):
    """ Get approved and rejected email lists from current app state. """
    emails = []
    include = []
    exclude = []
    
    if (app.filepath):
        emails = extract_email_lst(app)

    # Obtain filter lists from text boxes
    include = process_list(app.include_box.get("1.0", 'end-1c'))
    exclude = process_list(app.exclude_box.get("1.0", 'end-1c'))

    return filter_emails(emails, include, exclude)


 






        
        
        
                    
        
        

    
