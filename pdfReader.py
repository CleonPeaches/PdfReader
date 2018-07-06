import PyPDF2 as pydf, os, openpyxl, sys, datetime, re

now = datetime.datetime.now()

theNumber = 0.0

columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I',
           'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R']

class Employee(object):
    
    def __init__(self, name, cashRemitted, gcActivated, taxEx, transferIn, transferOut,
                 visa, mc, amex, disc, dinersCarte, comps, gcRedeem, tipshareMoney,
                 tipsharePercent, cashDue, overShort):
        self.name = name
        self.cashRemitted = cashRemitted
        self.gcActivated = gcActivated
        self.taxEx = taxEx
        self.transferIn = transferIn
        self.transferOut = transferOut
        self.visa = visa
        self.ms = mc
        self.amex = amex
        self.disc = disc
        self.dinersCarte = dinersCarte
        self.comps = comps
        self.gcRedeem = gcRedeem
        self.tipshareMoney = tipshareMoney
        self.tipsharePercent = tipsharePercent
        self.cashDue = cashDue
        self.overShort = overShort

def read_pdf(pageObject):
    
    allText = pageObject.extractText().split(' ')
    allTextList = list(allText)

    index = 0
    newList = []

    #Remove \n characters from each string and add to list.
    for i in allTextList:
        originalString = allTextList[index]
        if '\n' in originalString:
            newString = originalString.replace('\n', '')
            newList.append(newString)
        else: 
            newList.append(originalString)
        index += 1
        
    return newList


def default_excel_sheet():
    
    wb = openpyxl.Workbook()

    ws = wb.active

    standard_label = ['J. Alexander\'s / Redlands', 'Cash Out Spreadsheet', 'Denver',
                      now.strftime('%A - %b %d, %Y'), 'Version 12.07.2016']

    xAxisLabelOne = ['Cash', 'Gross', 'GC', 'Tax Ex', 'Transfer', 'Transfer', 'Diners/',
                   'GC', 'Tip', 'Tip', 'Cash', 'Over/']

    xAxisLabelTwo = ['Server/Pubkeep', 'Remitted', 'Sales', 'Activated', 'Sales', 'In', 'Out', 'Visa',
                     'MC', 'Amex', 'Disc', 'Carte', 'Comps', 'Redeem', 'Share $', 'Share %',
                     'Due', 'Short']
    
    
    #Populate header of Excel sheet with restaurant info, date and version.
    for i in range(0, 5):
        current_element = 'A' + str(i+1)
        ws[current_element] = standard_label[i]

    #Populate X and Y axes with labels
    for i in range(0, 6):
        current_element = columns[i + 1] + str(6)
        ws[current_element] = xAxisLabelOne[i]

    ws['L6'] = 'Diners/'

    for i in range(0, 5):
        current_element = columns[i + 13] + str(6)
        ws[current_element] = xAxisLabelOne[i + 7]

    for i in range(0, 17):
        current_element = columns[i] + str(7)
        ws[current_element] = xAxisLabelTwo[i]

    wb.save('example.xlsx')

    wbpath = os.path.dirname(os.path.realpath('example.xlsx'))

    print('Saved empty excel sheet to ' + wbpath + '.')

    return wb


def parse_float(theString):
    theNumber = re.findall(r'[-]?\d*\.\d+|\d+', theString)
    
    return theNumber


if __name__ == '__main__':
    
    #TODO If today is Monday, make new workbook. Else, populate next day.
    #TODO set curdir to desktop
    try:
        thePdf = open('C:\\Users\\drewe\\Desktop\\employee totals.pdf', 'rb')
    except IOError:
        print('An error occurred while reading the pdf file. Make sure the file is ' +
              'named correctly, and that it\'s saved to the proper location.')
        sys.exit()
        
    pdfObject = pydf.PdfFileReader(thePdf)
    print('Pdf object created.\n')
    numPages = pdfObject.getNumPages()
    print('Pages detected: ' + str(numPages) + '\n')
    employeeList = []

    #This loop creates an employee object for each sheet of the pdf
    #and adds it to a list.
    for i in range(0, 1):
        pageObject = pdfObject.getPage(i)
        employeeData = read_pdf(pageObject)
        #for i in enumerate(employeeData):
        #    print(i)
        if employeeData[27] == 'AM':
            name = 'AM PUB'
        else:
            name = employeeData[27] + ' ' + employeeData[28]
            cashRemitted = parse_float(employeeData[227])
            print(cashRemitted)

            gcActivated = ''
            taxEx = ''
            print(employeeData[55])
            transferredIn = parse_float(employeeData[55])
            
            
        print(name)
        
    
    
    empty_sheet = default_excel_sheet()
    
