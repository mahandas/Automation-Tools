from docx import Document
import os
import sys
from glob import glob
import re
import win32com.client as win32
from win32com.client import constants
import logging
import inspect
import time



def generate_log(log_path):
    function_name = inspect.stack()[1][3]
    logger = logging.getLogger(function_name)
    d = time.strftime("\n%d-%m-%Y   ") + time.strftime("%H:%M:%S")
    fh = logging.FileHandler(log_path.format(function_name))
    fh_format = logging.Formatter(str(d) + ' %(levelname)s %(message)s')
    fh.setFormatter(fh_format)
    logger.addHandler(fh)
    logger.setLevel(logging.DEBUG)
    return logger


def ConvertTodocx(doc, logger_object):

    # Extract the text from the DOCX file object infile and write it to 
    # a PDF file.

    try:
        if(".docx" not in doc):
            logger_object.debug("Convert2Doc function intialized")
            try:             
                word = win32.gencache.EnsureDispatch('Word.Application')
                word.Visible = False
                docer = word.Documents.OpenNoRepairDialog(doc, False, False)
                docer.Activate ()
            except Exception as e:
                time.sleep(10)
                word = win32.gencache.EnsureDispatch('Word.Application')
                word.Visible = False
                docer = word.Documents.OpenNoRepairDialog(doc, False, False)
                docer.Activate ()
            logger_object.info("new word initialised")
            new_file_abs = os.path.abspath(doc)
            new_file_abs = re.sub(r'\.\w+$', '.docx', new_file_abs)
            logger_object.info("new document created :- ")
            logger_object.info(new_file_abs)
            word.ActiveDocument.SaveAs(new_file_abs, FileFormat=constants.wdFormatXMLDocument)
            docer.Close(False)
            word.Application.Quit(-1)
        else:
            logger_object.info("Something went wrong in Convert2Doc")
        
    except Exception as e:
        logger_object.exception(str(e))



if __name__ == "__main__":       
    try:                   
            
        if len(sys.argv) == 3: #if length is 3 then convert format mode
            file1 = sys.argv[1]
            logpath = sys.argv[2]
            logger_object = generate_log(logpath)
            logger_object.info("PDF Merger Started - convert mode ")
            extn = file1.split(".")
            file1extn = extn[len(extn) - 1]
            if file1extn == "doc" :
                    logger_object.info("Converting file to docx")
                    ConvertTodocx(file1, logger_object)
                    logger_object.info("Conversion successful ")
            elif(file1extn == "docx"):
                logger_object.info("File is already in docx format.")
            else:
                logger_object.info("File is not a document. Please check the file type.")
        else :
            print("arguments not proper")
    except Exception as e:
            print(str(e))
