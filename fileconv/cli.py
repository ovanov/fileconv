"""
This script will convert .txt , .docx or .doc & .xlsx or .xls / .csv to .pdf
It uses 'fpdf', 'docx2pdf' & 'win32' to execute this task.

The programm crawls over one directory with (possible multiple) subdirectories and identifies files, by their extension and converts them to pdf, while leaving the old ones 


@Author: ovanov
@Date: 26.03.21
"""

import os
import argparse
import pathlib
from typing import Dict

from tqdm.std import tqdm


from .office_to_pdf import MS
from .txt_to_pdf import Conv
from .pdf_to_txt import Pdf

def argument_parser() -> Dict:
    parser = argparse.ArgumentParser('fileconv',description='Command line tool for file conversion to PDF. Supports MS Word, Excel and txt files.')

    parser.add_argument('--pdf',
    help='(required) filepath to the directory, in which the files should be converted to PDF',
    nargs='?',
    type=str,
    default=False)

    parser.add_argument('--txt',
    help='(required) filepath to the directory, in which the files should be converted to TXT',
    nargs='?',
    type=str,
    default=False)

    parser.add_argument('--output', '-o',
    help='(required) Give the path and the name of the output directory.',
    nargs='?',
    type=str,
    default=False)

    return parser


def crawler(p, output):
    """
    This function takes the path variable "p" and goes over all directories 
    """
    dir_len = len(next(os.walk(p))[1])
    for root, dirs, files in tqdm(os.walk(p), total=dir_len):
        #os.walk yields a 3-tuple of strings, wich can be concatenated
        if len(root) != 0:
            # maybe the directory is empty or yields no files (empty list), we should make sure that there is no error
            for filename in files:

                ext = pathlib.Path(filename).suffix # get file extension

                if ext == ".docx" or ext == ".doc":
                    file = os.path.join(root, filename)
                    MS.word_to_pdf(file, ext, filename, output)

                elif ext == ".xlsx" or ext == ".xls":
                    file = os.path.join(root, filename)
                    MS.excel_to_pdf(file, ext, filename, output)

                elif ext == ".txt":
                    file = os.path.join(root, filename)
                    Conv.txt_to_pdf(file, filename, output) #TODO add the txt class
                
                else:
                    pass
    
        else:
            raise KeyError("The directory is empty")
    return

def pdf_crawler(p, output):
    """
    This function takes the path variable "p" and goes over all directories 
    """
    dir_len = len(next(os.walk(p))[1])
    for root, dirs, files in tqdm(os.walk(p), total=dir_len):
        #os.walk yields a 3-tuple of strings, wich can be concatenated
        if len(root) != 0:
            # maybe the directory is empty or yields no files (empty list), we should make sure that there is no error
            for filename in files:

                ext = pathlib.Path(filename).suffix # get file extension

                if ext == ".pdf":
                    file = os.path.join(root, filename)
                    Pdf.to_text(file, ext, filename, output)
                
                else:
                    pass
    
        else:
            raise KeyError("The directory is empty")
    return

def main():
    """
    Change the "path" variable, in order to pass the directory
    """
    # get argument parser
    parser = argument_parser()
    args = parser.parse_args()
    args_dict = {
        arg: value for arg, value in vars(args).items()
        if value is not None
    }

    if len(args_dict['pdf']) < 1 and len(args['txt']) < 1: 
        raise KeyError('Please give a path to a file or a dirctory path.')
    if len(args_dict['pdf']) == 1 and len(args['txt']) == 1: 
        raise KeyError("Please give only one task. fileconv can't parse in both ways at once.")

    if len(args_dict['pdf']) == 1:
        path = args_dict['pdf']
        output = args_dict['output']
        crawler(path, output)

    else:
        path = args_dict['txt']
        output = args_dict['output']
        pdf_crawler(path, output)

    return

if __name__ == "__main__":
    main()
