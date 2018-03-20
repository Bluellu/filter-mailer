# Excel manipulation functions
from tkinter import *
from tkinter.filedialog import askopenfilename
from os.path import normpath, basename

import xlrd

def get_filepath(app, label):
    app.filepath = askopenfilename(initialdir = "C:/",
                                title = "import file",
                                filetypes =
                                [('Excel', ('*.xls', '*.xlsx')),
                                ('CSV', '*.csv',)])
    
    # Update displayed file name
    if (app.filepath):
        label.config(text = basename(normpath(app.filepath)))
        print(app.filepath)

