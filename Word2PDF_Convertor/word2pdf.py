import sys
import os
import comtypes.client

wdFormatPDF = 17


word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open('C:\Users\mahan.das\AppData\Local\Programs\Python\Python36\test1.docx')
doc.SaveAs("C:\Users\mahan.das\AppData\Local\Programs\Python\Python36\ch.pdf"
, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()
