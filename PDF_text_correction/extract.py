from pdfminer.high_level import extract_text
from fpdf import FPDF
from PyPDF2 import PdfFileReader, PdfFileWriter
from tkinter import ttk, filedialog
import tkinter as tk
import os

class Extract(tk.Tk):
    def __init__(self):
        super().__init__()

        #GUI Base
        self.title("Text Extraction")
        self.intro = ttk.Frame(self, padding = "10")
        self.intro.grid(column = 0, row = 0)
        self.fileselect = ttk.Frame(self, padding = "10")
        self.fileselect.grid(column = 0, row = 1)
        self.extract = ttk.Frame(self, padding = "10")
        self.extract.grid(column = 0, row = 2)
        self.save = ttk.Frame(self)
        self.save.grid(column = 0, row = 3)

        #File Selection
        filetext = ttk.Label(self.fileselect, text = "Select file to view text")
        filetext.grid(column = 0, row = 0)
        self.filebox = ttk.Label(self.fileselect, text = "")
        self.filebox.config(width = 50, background = "white", borderwidth = 3, relief = "solid")
        self.filebox.grid(column = 0, row = 1)

        #Text Display
        self.textbox = tk.Text(self.extract, width = 50, height = 10)
        self.textbox.grid(column = 0, row = 0)


        #Buttons
        filebtn = ttk.Button(self.fileselect, text = "Browse", command = self.openfile)
        filebtn.grid(column = 1, row = 1)
        extract_btn = ttk.Button(self.fileselect, text = "Extract Text", command = self.extracttext)
        extract_btn.grid(column = 0, row = 2)
        savenew_btn = ttk.Button(self.save, text = "Save text as new file", command = self.savenew)
        savenew_btn.grid(column = 0, row = 1)
        append_btn = ttk.Button(self.save, text = "Append text to PDF", command = self.append)
        append_btn.grid(column = 1, row = 1)

    def openfile(self):
        self.open_file = filedialog.askopenfilename(filetypes=[("pdf files", "*.pdf")])
        self.filebox.config(text = self.open_file)

    def extracttext(self):
        text = extract_text(self.open_file)
        self.textbox.insert("1.0", text)

    def savenew(self):
        saveas = self.open_file.replace("pdf", "txt")
        newtxt = self.textbox.get("1.0", "end")
        with open(saveas, "w") as f:
            f.write(newtxt)
            

    def append(self):
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Times", size = 10)
        newtxt = self.textbox.get("1.0", "end")
        newname = self.open_file.replace(".pdf", "_edit.pdf")
        print(newname)
        pdf.multi_cell(180, 5, txt = newtxt)
        pdf.output(newname)

        input_ = PdfFileReader(open(self.open_file, "rb"))
        pagesi = input_.getNumPages()
        newtext = PdfFileReader(open(newname, "rb"))
        pageso = newtext.getNumPages()
        output = PdfFileWriter()
        for i in range(pagesi):
            output.addPage(input_.getPage(i))
        for i in range(pageso):
            output.addPage(newtext.getPage(i))
        outputStream = open(newname, "wb")
        output.write(outputStream)
        outputStream.close()

if __name__ == "__main__":
    app = Extract()
    app.mainloop()