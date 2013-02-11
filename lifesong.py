# Based on the accepted StackOverflow answer at http://stackoverflow.com/questions/6011115/doc-to-pdf-using-python
import sys
import os
import comtypes.client

wdFormatPDF = 17

in_file = os.path.abspath(sys.argv[1])
out_file = os.path.abspath(sys.argv[2])

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_file, FileFormat=wdFormatPDF)
doc.Close()
word.Quit()
