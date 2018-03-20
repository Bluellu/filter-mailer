import tkinter as tk
import excel_manipulation as em

def main():
    root = tk.Tk()
    root.title("Filter Mailer")
    root.geometry("400x200")
    app = MainApp(root)

    #Start event loop
    root.mainloop()
    

class MainApp:
    def __init__(self, parent):
        self.parent = parent
        self.frame = tk.Frame(self.parent)
        self.frame.pack(fill = 'both', expand = 1)
        self.filepath = ""
        
        #Labels
        title_lbl = tk.Label(self.frame, text = "Filter Mailer")
        file_lbl = tk.Label(self.frame, text = "None selected ")

        #Button for selecting an excel filepath
        import_bttn = tk.Button(self.frame, text = "Select File",
                        fg = "white",
                        bg = "navy",
                        command = (lambda: em.get_filepath(self, file_lbl)))

        #Add items to grid
        title_lbl.grid(column = 0, row = 0, columnspan = 3, sticky = "nsew")
        import_bttn.grid(column = 0, row = 1, sticky = "nsew")
        file_lbl.grid(column = 1, row = 1, columnspan = 2, sticky = "w")
        
        #Adjust column and row expansion
        self.frame.grid_rowconfigure(2, weight = 1)
        self.frame.grid_columnconfigure(2, weight = 1)
        

if __name__ == "__main__":
    main()
