from pdfminer.high_level import extract_text
from fpdf import FPDF
from googletrans import Translator
import re
import sys

translator = Translator()
pdf = FPDF()

def main():
    try:
        if len(sys.argv) < 3:
            sys.exit("Too few command line arguments. Format must be: python.exe PDF_translator.py inputfile.pdf outputfile.pdf")
        elif len(sys.argv) > 3:
            sys.exit("Too many command line arguments. Format must be: python.exe PDF_translator.py inputfile.pdf outputfile.pdf")
        elif sys.argv[1].endswith(".pdf") == False:
            sys.exit("Input not a valid pdf")
        elif sys.argv[2].endswith("pdf") == False:
            sys.exit("Output not a valid pdf")
        else:
            text = extract_text(sys.argv[1]).strip()
            if len(text) > 5000:
                sys.exit("Input pdf is too long!")
            else:
                trans = translation(text)
                pdf_create(trans)
    except FileNotFoundError:
        sys.exit("Can't find input file")

def translation(x):
    #RegEx to handle the code left by Translator at the start and end of the text
    prefix = r"^Translated\(src.+text="
    suffix = r"pronunciation.+\"\)$"

    #Get text from pdf, translate to English

    translated = str(translator.translate(x))

    #Remove code left by Translator at the start and end of the text
    prefix_removed = re.sub(prefix, "", translated)
    suffix_removed = re.sub(suffix, "", prefix_removed)
    return suffix_removed

def pdf_create(newtext):
    #Create pdf in English
    pdf.add_page()
    pdf.set_font("Times", size = 10)
    pdf.multi_cell(180, 5, txt = newtext)
    pdf.output(sys.argv[2])

if __name__ == "__main__":
    main()
