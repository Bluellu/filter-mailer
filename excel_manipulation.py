# Excel manipulation functions
from collections import OrderedDict
import pandas as pd     
        
def extract_email_lst(app):
    """ Extract and return a list of emails from an excel file column.

        Args:
            app: The app instance containing the desired excel filepath.
        Returns:
            A single list of email strings.
            
    """
    ordered_email_lst = []
    if (app.filepath):
        col_name = 'E-Mail'
        
        df = pd.read_excel(app.filepath)
     
        if col_name in df.columns.values:
            email_lst = df[df[col_name].notnull()][col_name].values #TODO: Allow different spellings with regex

            #Remove duplicates
            ordered_email_lst = list(OrderedDict.fromkeys(email_lst))
        
    return ordered_email_lst

    #TODO: allow user to select correct email column if multiple available


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
    #Remove all spaces
    no_spaces = "".join(string.split())

    #Split at commas
    split_lst = no_spaces.split(",")

    #Remove empty strings and return
    return list(filter(None, split_lst))
    
def get_sorted_emails(app):
    """ Get approved and rejected email lists from current app state. """
    emails = []
    include = []
    exclude = []
    
    if (app.filepath):
        emails = extract_email_lst(app)

    #Obtain filter lists from text boxes
    include = process_list(app.include_box.get("1.0", 'end-1c'))
    exclude = process_list(app.exclude_box.get("1.0", 'end-1c'))

    return filter_emails(emails, include, exclude)




        
        
        
                    
        
        

    
