# pdfReader
This is a script written for the management management of J. Alexander's in Denver, meant to expedite the cashout process.

It takes the nightly cashout PDFs, extracts the text for each employee, uses regular expressions to correctly parse the dollar amounts,
and turns it into an Excel sheet.

The functionality of the PyPDF2 module is highly inconsistent when extracting text from a PDF; this code will not work if the format of 
the PDFs is changed even slightly.
