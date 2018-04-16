import tkinter as tk
import tkinter.scrolledtext as tkscrolled
import ui_operations as uop

def main():
    root = tk.Tk()
    root.title("Filter Mailer")
    root.geometry("540x400")
    app = MainApp(root)

    #Start event loop
    root.mainloop()
    

class MainApp:
    def __init__(self, parent):
        self.root = parent
        self.frame = tk.Frame(self.root)
        self.frame.pack(fill = 'both', expand = 1)
        self.filepath = ""

        #Styling parameters
        box_width = 26
        box_height = 10

        #Button for selecting an excel filepath
        import_bttn = tk.Button(self.frame, text = "Select File", width = 0,
                        fg = "white",
                        bg = "navy",
                        command = (lambda: uop.get_filepath(self, file_lbl)))
        #Label for selected file
        file_lbl = tk.Label(self.frame, text = "None selected ")

        #Filter keyword boxes
        filter_lbl = tk.Label(self.frame, text = "Filter keywords (separate with commas): ")
        
        include_lbl = tk.Label(self.frame, text = "Include", fg = "navy", bg = "light blue")
        self.include_box = tkscrolled.ScrolledText(self.frame, width = box_width, height = box_height)
        
        exclude_lbl = tk.Label(self.frame, text = "Exclude", fg = "navy", bg = "light blue")
        self.exclude_box = tkscrolled.ScrolledText(self.frame, width = box_width, height = box_height)

        #Button for previewing filtered email list
        preview_bttn = tk.Button(self.frame, text = "Preview",
                        fg = "white",
                        bg = "navy",
                        command = (lambda: uop.filter_preview(self)))

        #Button for opening email creation
        email_bttn = tk.Button(self.frame, text = "Proceed to email creation",
                        fg = "white",
                        bg = "navy",
                        command = (lambda: uop.email_creation(self)))        

        #Add items to grid
        import_bttn.grid(column = 0, row = 0, columnspan = 1, padx = 5, pady = 15, sticky = "new")
        
        file_lbl.grid(column = 1, row = 0, columnspan = 3, sticky = "w")

        filter_lbl.grid(column = 0, row = 1, columnspan = 2, padx = 5, sticky = "w")

        include_lbl.grid(column = 0, row = 2, columnspan = 2, padx = 5, sticky = "nsew")
        self.include_box.grid(column = 0, row = 3, columnspan = 2, padx = 5)

        exclude_lbl.grid(column = 2, row = 2, columnspan = 2, padx = 5, sticky = "nsew")
        self.exclude_box.grid(column = 2, row = 3, columnspan = 2, padx = 5)

        preview_bttn.grid(column = 3, row = 1, padx = 5, pady = 15, sticky = "nsew")

        email_bttn.grid(column = 1, row = 4, columnspan = 3, padx = (50, 15), pady = (30, 15), sticky = "nsew")
        
        #Adjust column and row configuration
        self.frame.grid(padx = 25)
        self.frame.grid_columnconfigure(0, weight = 1)
        self.frame.grid_columnconfigure(1, weight = 2)
        self.frame.grid_columnconfigure(2, weight = 2)
        self.frame.grid_columnconfigure(3, weight = 1)

        

if __name__ == "__main__":
    main()
