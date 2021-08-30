"""
This class will convert .pdf's back to .txt
It uses 'PyPDF2' to execute this task.


@Author: ovanov
@Date: 30.08.21
"""

import os

import PyPDF2


class Pdf():

    @staticmethod
    def to_text(text, *args):
        """
        converts pdf to txt and uses filepath from converter
        """
        with open(text, "rb") as f:
            #create reader variable that will read the pdf_obj
            pdfreader = PyPDF2.PdfFileReader(f)
            #This will store the number of pages of this pdf file
            x = pdfreader.numPages
            #create a variable that will select the selected number of pages
            pageobj = pdfreader.getPage(x+1)
            #(x+1) because python indentation starts with 0.
            #create text variable which will store all text datafrom pdf file
            txt = pageobj.extractText()

        file_loc_and_name = str(os.path.join(args[1], args[0][:-4])) + ".pdf" 
        with open(file_loc_and_name, "w") as T:
            T.writelines(txt)