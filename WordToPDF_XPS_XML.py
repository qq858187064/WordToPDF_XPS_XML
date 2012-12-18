# Word to PDF converter
import sys,os
import comtypes.client

wdFormatPDF = 17
wdFormatXPS = 18
wdFormatFlatXML = 19

#source word file
in_file = os.path.abspath(sys.argv[1])

#output PDF file
out_pdf = os.path.abspath(sys.argv[2])

#output XPS file
out_xps = os.path.abspath(sys.argv[3])

#output as Flat XML file
out_xml = os.path.abspath(sys.argv[4]) 

word = comtypes.client.CreateObject('Word.Application')
doc = word.Documents.Open(in_file)
doc.SaveAs(out_pdf, FileFormat=wdFormatPDF)
doc.SaveAs(out_xps, FileFormat=wdFormatXPS)
doc.SaveAs(out_xml, FileFormat=wdFormatFlatXML)
doc.Close()
word.Quit()
