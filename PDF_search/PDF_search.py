from pdfminer.high_level import extract_text
import tkinter as tk
from tkinter import EXTENDED, ttk, filedialog
import re, os, docx, winshell

class Search(tk.Tk):
    def __init__(self):
        super().__init__()

        #GUI Base
        self.title("Keyword Search")
        self.intro = ttk.Frame(self, padding = "10")
        self.intro.grid(column = 0, row = 0)
        self.foldersearch = ttk.Frame(self, padding = "10")
        self.foldersearch.grid(column = 0, row = 1)
        self.keywordsearch = ttk.Frame(self)
        self.keywordsearch.grid(column = 0, row = 2)
        self.submit = ttk.Frame(self)
        self.submit.grid(column = 0, row = 3)

        #Intro and explanation
        introtext = ttk.Label(self.intro, text = "Use this programme to search for a keyword (or key phrase) in pdf and Word files.")
        introtext.config(wraplength= 400)
        introtext.grid(column = 0, row = 0)

        #Folder Selection
        fldrtext = ttk.Label(self.foldersearch, text = "Select folder to search in:")
        fldrtext.grid(column = 0, row = 0)
        self.folderbox = ttk.Label(self.foldersearch, text = "")
        self.folderbox.config(width = 50, background = "white", borderwidth = 3, relief = "solid")
        self.folderbox.grid(column = 0, row = 1)
        folderbtn = ttk.Button(self.foldersearch, text = "Browse", command = self.openfolder)
        folderbtn.grid(column = 1, row = 1)

        #Keyword Input
        kwordtext = ttk.Label(self.keywordsearch, text = "Type keyword or key phrase to search for:")
        kwordtext.grid(column = 0, row = 0)
        self.kword = tk.StringVar()
        self.keyword_entry = ttk.Entry(self.keywordsearch, textvariable = self.kword)
        self.keyword_entry.grid(column = 0, row = 1)

        #Submit
        self.submitbtn = ttk.Button(self.submit, text = "Search", command = self.search)
        self.submitbtn.grid(column = 0, row = 0)

    def openfolder(self):
        #Dialog box to choose folder to search
        self.open_folder = filedialog.askdirectory()
        self.folderbox.config(text = self.open_folder)

    def search(self):
        #Prompt user to input a keyword if they haven't
        if self.kword.get() == "":
            self.error("keyword")
        else:
            try:
                self.lst = []
                folder = os.listdir(self.open_folder)
                #Extract text from pdf and Word files to search for keyword
                for filename in folder:
                    file = self.open_folder + "\\" + filename
                    if filename.endswith(".pdf"):
                        f = extract_text(file)
                        self.keyword(filename, f)
                    elif filename.endswith(".doc") or file.endswith(".docx"):
                        doc = docx.Document(file)
                        texts = []
                        for para in doc.paragraphs:
                            texts.append(para.text)
                        fulltext = " ".join(texts)
                        self.keyword(filename, fulltext)
                #If keyword not found in any files
                if len(self.lst) == 0:
                    self.lst.append("No files found!")
                
                self.newwindow()
            except AttributeError:
                #Prompts user to choose a folder is they haven't
                self.error("folder")

    def error(self, label):
        #Message to prompt user to input folder or keyword if they haven't
        wind = tk.Toplevel(self)
        self.title("Error!")
        self.frame = ttk.Frame(wind, padding = "10")
        self.frame.grid(column = 0, row = 0)
        aerrorlbl = ttk.Label(self.frame, text = f"Please choose a {label} to search!")
        aerrorlbl.grid(column = 0, row = 0)

    def keyword(self, filename, txt):
        #Search extracted text for selected keyword
        search = re.search(self.kword.get(), txt, re.IGNORECASE)
        if search:
            self.lst.append(filename)


    def newwindow(self):
        #Open new window with list of files containing keyword
        wind = tk.Toplevel(self)
        self.title("Files Containing Keyword")
        self.boxframe = ttk.Frame(wind, padding = "10")
        self.boxframe.grid(column = 0, row = 0)
        self.buttonsframe = ttk.Frame(wind, padding = "10")
        self.buttonsframe.grid(column = 1, row = 0)

        boxlbl = ttk.Label(self.boxframe, text = "Files containing keyword:")
        boxlbl.grid(column = 0, row = 0)
        filelist = tk.StringVar(value = self.lst)
        self.box = tk.Listbox(self.boxframe, listvariable = filelist, selectmode = EXTENDED, width = 50)
        self.box.grid(column = 0, row = 1)

        buttonstext = ttk.Label(self.buttonsframe, text = "Select files you are interested in. \n These can be opened externally using the button below. \n You can also save shortcuts to the selected files in a new subfolder, for easy perusal at a later time.")
        buttonstext.grid(column = 0, row = 0)
        openbtn = ttk.Button(self.buttonsframe, text = "Open Files", command = self.openfiles)
        openbtn.grid(column = 0, row = 1)
        savebtn = ttk.Button(self.buttonsframe, text = "Save", command = self.savefiles)
        savebtn.grid(column = 0, row = 2)

    def openfiles(self):
        #Opens selected files
        files = self.box.curselection()
        for i in files:
            f = os.path.join(self.open_folder, self.box.get(i))
            os.startfile(f)

    def savefiles(self):
        #Saves shortcuts to selected files in new folder
        foldername = "Keyword - " + f"'{self.kword.get()}'"
        source = os.path.join(self.open_folder, foldername)
        try:
            os.mkdir(source)
        except FileExistsError:
            pass

        files = self.box.curselection()
        for i in files:
            target = os.path.join(self.open_folder, self.box.get(i))
            shortcutname = self.box.get(i)[0:-5] + ".lnk"
            newpath = os.path.join(source, shortcutname)
            shortcut = winshell.shortcut(target)
            shortcut.working_directory = source
            shortcut.write(newpath)
            shortcut.dump()
            

if __name__ == "__main__":
    app = Search()
    app.mainloop()






