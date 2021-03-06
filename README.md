<h1 align="center">fileconv</h1>
<p align="center">Convert your files to PDF's</p>

<p align="center">
	<a href="https://github.com/ovanov/fileconv#ovanov"><img src="https://img.shields.io/github/languages/code-size/ovanov/fileconv?color=greem&label=package%20size" height="20"/></a>
    <a href="https://github.com/ovanov/fileconv#ovanov"><img src="https://img.shields.io/github/license/ovanov/fileconv?color=black" height="20"/></a>
    <a href="https://pypi.org/project/fileconv"><img src="https://img.shields.io/pypi/pyversions/fileconv">
    <a href="https://github.com/ovanov/fileconv/blob/main/setup.py"><img src="https://img.shields.io/pypi/wheel/fileconv?color=yellow">
</p>

<p align="center"><a href="https://github.com/ovanov/fileconv#ovanov"><img src="https://github.com/ovanov/gifs/blob/main/filconvdemo.gif" width="130%"/></a></p><br/>


## :computer: How does it work?

**fileconv** gives you the possibility to conveniently convert your *text*, MS *Word* or MS *Excel* files to PDF-A2 standard, by using a simple and intuitive command line interface.


This program was made to convert digital born documents to PDF files, in order to have a long term archiving solution for all office documents. It is a free and open source alternative to proprietary converters and it is taking advantage of the extremely powerful [Pywin32](https://github.com/mhammond/pywin32) and the [pyfpdf](https://github.com/reingart/pyfpdf) packages. 

This Project has been brought to life with the help of the [AfZ](https://www.afz.ethz.ch/) (Archive of Contemporary History) at [ETH Zürich](https://ethz.ch/en.html).

## Overview

The fileconv tool is mainly used as a fast solution for file conversion via the command line. The *Conversion* package of this library is also available for in depth use in your python code. In order to use its methods, simply import the fileconv package.

At the moment, the package supports files that are
- WORD
    - .doc
    - .docx
- EXCEL
    - .xls
    - .xlsx
- TEXT
    - .txt

Support for more file types will be addes in the future.

The package is now also able to convert PDF files back to .txt files.

## Guide

The following shows how to get and use **fileconv**.

### Installation

    $ pip install fileconv

Consider that you might have to add the installation folder to your `PATH`.

If you would rather like to customize the code to your needs, grab a stable version under "Releases". All the files are extensively commented as well, in order to make the files more user friendly.

### Usage

When in a terminal, specify:

    $ fileconv --pdf path/to/dir --output path/to/output/dir

The program takes a directory, which is populated with **at least one** (!) file or subdirectory and takes an output location as a positional argument. It doesn't matter if your files are in a nested structure as shown in the example below:

    input_dir
    │
    ├── xy.xlsx
    ├── my_dir
    │   ├── xy.doc
    │   ├── ab.txt 
    │   └── sub_dir
    │        └── cd.docx
    └── ...


The program 'walks' through any directory structure and places all the files in a single folder that was specified in the `--output` flag.

Depending on your directory structure, the amount of files and your CPU, the process can take some time. With 1000 directories, 15'000 files to process, from which 800 are MS Office conversions, the program needs about 2 minutes and 30 seconds to finish. To not let you worry, we have implemented a progress bar that looks like this while running:

    5%|██████▌                                                             | 55/1003 [00:08<01:23, 11.38it/s]

This tool was based on the excellent [tqdm](https://github.com/tqdm/tqdm) library.

### Convert PDF files back to text

In order to convert PDF's back to text, use the `--txt` flag as shown bellow:

    $ fileconv --txt path/to/dir --output path/to/output/dir
	    
<p align="center"><a href="https://github.com/ovanov/fileconv#ovanov"><img src="https://github.com/ovanov/gifs/blob/main/filconvdemo_txt.gif" width="100%"/></a></p><br/>

Note, that the larger the files are, the longer the conversion takes.


### :bug: Common errors

**1. interrupting the programm**

The program has one weakness regarding its running time. The pywin32 library has to 'open' the according MS programs in the background in order to properly convert the files. This process is prone to an error, *if the program is interrupted while running*, i.e. by pressing **CTRL + C**. In this case, fileconv will prompt an error the next time it converts a MS file type, because the background process has not been closed. To prevent this, one can open a *terminal* and type the following commands:

    $ python

or (depending on your OS and python versions) 

    $ python3

this will open the python interpreter

    >>> import win32com.client
    >>> win32com.client.Dispatch('Word.Application').Close()
    >>> win32com.client.Dispatch('Excel.Application').Close()

This should prevent the error from occurring again.

**2. cross platform compatibility**

Even though this tool is intended and written to work on all operating systems, the [pywin32](https://github.com/mhammond/pywin32) library seems only to be reliable on Mircrosoft Windows. UNIX based file systems, like MacOS and Linux tend to prompt errors during the library's installation. As referred in this [issue thread](https://github.com/malwaredllc/byob/issues/19) the library is mostly developed on Windows and therefore causing instability when used on other systems. The library is probably working on fixes and is potentially fixing this error in future versions, which will be included in this package as well.

We are also looking for stable and reliable alternatives, but right now [pywin32](https://github.com/mhammond/pywin32) seems to be the best option.
