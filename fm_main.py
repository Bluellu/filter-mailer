import tkinter as tk
import excel_manipulation as em

def main():
    root = tk.Tk()
    root.title ("Filter Mailer")
    root.geometry("400x200")
    app = MainApp(root)

    #Start event loop
    root.mainloop()
    

class MainApp:
    def __init__(self, parent):
        self.parent = parent
        self.frame = tk.Frame(self.parent)
        self.filepath = ""

        label = tk.Label(self.frame, text = "None selected ")
        label.pack(side = tk.LEFT)

        #Button for importing an excel file
        self.import_bttn = tk.Button(self.frame, text = "Select File",
                        fg = "white",
                        bg = "navy",
                        command = (lambda: em.get_filepath(self, label)))

        self.import_bttn.pack(side = tk.RIGHT)
        self.frame.pack()
        

if __name__ == "__main__":
    main()
