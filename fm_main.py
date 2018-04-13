import tkinter as tk
import tkinter.scrolledtext as tkscrolled
import excel_manipulation as em

def main():
    root = tk.Tk()
    root.title("Filter Mailer")
    root.geometry("600x300")
    app = MainApp(root)

    #Start event loop
    root.mainloop()
    

class MainApp:
    def __init__(self, parent):
        self.root = parent
        self.frame = tk.Frame(self.root)
        self.frame.pack(fill = 'both', expand = 1)
        self.filepath = ""

        #Button for selecting an excel filepath
        import_bttn = tk.Button(self.frame, text = "Select File",
                        fg = "white",
                        bg = "navy",
                        command = (lambda: em.get_filepath(self, file_lbl)))
        #Label for selected file
        file_lbl = tk.Label(self.frame, text = "None selected ")

        #Text input widgets for filter words
        incld_lbl = tk.Label(self.frame, text = "Include", fg = "navy", bg = "light blue")
        incld_box = tkscrolled.ScrolledText(self.frame, width = 20, height = 10)
        
        excld_lbl = tk.Label(self.frame, text = "Exclude", fg = "navy", bg = "light blue")
        excld_box = tkscrolled.ScrolledText(self.frame, width = 20, height = 10)

        #Add items to grid
        import_bttn.grid(column = 0, row = 0, padx = 10, pady = 10, sticky = "nsew")
        file_lbl.grid(column = 1, row = 0, columnspan = 2, sticky = "w")

        incld_lbl.grid(column = 1, row = 2, padx = 5, sticky = "nsew")
        incld_box.grid(column = 1, row = 3, padx = 5, columnspan = 1, rowspan = 2)

        excld_lbl.grid(column = 2, row = 2, padx = 5, sticky = "nsew")
        excld_box.grid(column = 2, row = 3, padx = 5, columnspan = 1, rowspan = 2)
        
        #exclude_box.grid(column = 2, row = 2, columnspan = 1, rowspan = 2, sticky = "nsew")     
        
        #Adjust column and row expansion
       # self.frame.grid_rowconfigure(2, weight = 2)
       # self.frame.grid_columnconfigure(2, weight = 2)

        

if __name__ == "__main__":
    main()
