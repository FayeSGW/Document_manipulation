from pdfminer.high_level import extract_text
import re
import os
import sys
import docx

def main(): 

    f_input = input("Folder to search: ").replace("/", "\\")
    folder = os.listdir(f_input)

    global text
    text = input("What keyword do you want to find? ")
    global lst
    lst = []

    if len(lst) == 0:
        print("Keyword not found anywhere!")
    else:
        print(lst)

    for filename in folder:
        if f_input.endswith("\\"):
            file = f_input + filename
        else:
            file = f_input + "\\" + filename
        if file.endswith(".pdf"):
            f = extract_text(file)
            keyword(filename, f)
        elif file.endswith(".doc") or file.endswith(".docx"):
            doc = docx.Document(file)
            texts = []
            for para in doc.paragraphs:
                texts.append(para.text)
            fulltext = " ".join(texts)
            keyword(filename, fulltext)

    choice = input("Open files? ").lower().strip()
    if choice == "yes" or choice == "y":
        for i in lst:
            os.startfile(i)
    else:
        sys.exit()

def keyword(filename, txt):
    search = re.search(text, txt, re.IGNORECASE)
    if search:
        lst.append(filename)
    else:
        return "no"

if __name__ == "__main__":
    main()






