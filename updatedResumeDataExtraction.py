import os, glob, re
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
from docx import Document
import PyPDF2


def convertDocxToText(path):
    txt = []
    document = Document(path)
    txt.append("\n".join([para.text for para in document.paragraphs]))
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                txt.append("\n".join([para.text for para in cell.paragraphs]))

    return ('\n'.join([str(e) for e in txt]))



def convertPDFToText(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    password = ""
    maxpages = 0
    caching = True
    pagenos = set()
    for page in PDFPage.get_pages(fp, pagenos, maxpages=maxpages, password=password, caching=caching,
                                  check_extractable=True):
        interpreter.process_page(page)
    fp.close()
    device.close()
    string = retstr.getvalue()
    retstr.close()
    string = string.replace('\n','')
    return string

def convertPDFToTextUsingPypdf2(path):
    pdfFileObj = open(path, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
    num_pages = pdfReader.numPages
    count = 0
    text = ""

    while count < num_pages:
        pageObj = pdfReader.getPage(count)
        count += 1
        text += pageObj.extractText()
    return text

class exportToCSV:
    def __init__(self, fileName='resultsCSV.csv', resetFile=False):
        headers = [
            'EMPLOYEE NAME',
            'DATE OF BIRTH',
            'GENDER',
            'NATIONALITY',
            'CURRENT ADDRESS'
        ]
        if not os.path.isfile(fileName) or resetFile:
            # Will create/reset the file as per the evaluation of above condition
            fOut = open(fileName, 'w')
            fOut.close()
        fIn = open(fileName)  ########### Open file if file already present
        inString = fIn.read()
        fIn.close()
        if len(inString) <= 0:  ######### If File already exsists but is empty, it adds the header
            fOut = open(fileName, 'w')
            fOut.write(','.join(headers) + '\n')
            fOut.close()

    def write(self, infoDict):
        fOut = open('resultsCSV.csv', 'a+')
        # Individual elements are dictionaries
        writeString = ''
        try:
            writeString += str(infoDict['EMPLOYEE NAME']) + ','
            writeString += str(infoDict['DATE OF BIRTH']) + ","
            writeString += str(infoDict['GENDER']) + ","
            writeString += str(infoDict['NATIONALITY']) + ","
            writeString += str(infoDict['CURRENT ADDRESS']) + ",\n"
            fOut.write(writeString)
        except:
            fOut.write('FAILED_TO_WRITE\n')
        fOut.close()


class Parse():
    information = []
    inputString = ''

    def __init__(self):
        print('Starting Program...')

        # Glob module matches certain patterns
        doc_files = glob.glob("resumes/*.doc")
        docx_files = glob.glob("resumes/*.docx")
        pdf_files = glob.glob("resumes/*.pdf")

        files = set(doc_files + docx_files + pdf_files)
        files = list(files)
        print("%d files identified" % len(files))

        for file in files:
            print("Reading File %s" % file)
            # info is a dictionary that stores all the data obtained from parsing
            info = {}

            self.inputString = self.readFile(file)
            info['Filename'] = file

            self.getName(self.inputString, info, file)
            self.getDOB(self.inputString, info, file)
            self.getGender(self.inputString, info, file)
            self.getNationality(self.inputString, info, file)
            self.getCurrentAddress(self.inputString, info, file)

            csv = exportToCSV()
            csv.write(info)
            self.information.append(info)
            print(info)

    def readFile(self, fileName):

        extension = fileName.split(".")[-1]
        if extension == "pdf":
            return convertPDFToText(fileName)
        elif extension == "docx":
            try:
                return convertDocxToText(fileName)
            except:
                return ''
                pass
        else:
            print('Unsupported format')
            return ''

    #Function to get Employee Name.
    def getName(self, inputString, infoDict, file):  # incomplete function
        infoDict['EMPLOYEE NAME'] = 'NAME'

    #Function to get Date Of birth.
    def getDOB(self, inputString, infoDict, file):
        reg = re.search(r'(((DOB\s*:)|(D.O.B\s*:)|(date of birth\s*:)|(birthday\s*:))[- ]\s*\d{1,2}[-,./ ]*\s*((\d{1,2})|([a-zA-z]*))[-,./ ]\s*\d{4})|((DOB\s*:)|(D.O.B\s*:)|(date of birth\s*:)|(birthday\s*:))-*\s*(\d{1,2}((nd)|(th)|(rd))[-,./ ]*\s*[a-zA-z]*[-,./ ]*\s*\d{4})|((DOB)|(D.O.B)|(Date of Birth)|(birthday))\s*:-*\s*\d{1,2}[-,./ ]\s*\d{1,2}[-,./ ]\s*\d{4}|((DOB)|(D.O.B)|(Date of Birth)|(birthday))\s*:-*\s*[a-zA-Z]*[-,./ ]\d{1,2}[-,./ ]\s*\d{4}', inputString, re.IGNORECASE)
        ext = file[-4:]
        if reg is None and ext == '.pdf':
            text = convertPDFToTextUsingPypdf2(file)  # checking for text using PyPDF2 if not found using pdfminer
            text = text.replace('\n', '')
            reg = re.search(r'(((DOB\s*:)|(D.O.B\s*:)|(date of birth\s*:)|(birthday\s*:))[- ]\s*\d{1,2}[-,./ ]*\s*((\d{1,2})|([a-zA-z]*))[-,./ ]\s*\d{4})|((DOB\s*:)|(D.O.B\s*:)|(date of birth\s*:)|(birthday\s*:))-*\s*(\d{1,2}((nd)|(th)|(rd))[-,./ ]*\s*[a-zA-z]*[-,./ ]*\s*\d{4})|((DOB)|(D.O.B)|(Date of Birth)|(birthday))\s*:-*\s*\d{1,2}[-,./ ]\s*\d{1,2}[-,./ ]\s*\d{4}|((DOB)|(D.O.B)|(Date of Birth)|(birthday))\s*:-*\s*[a-zA-Z]*[-,./ ]\d{1,2}[-,./ ]\s*\d{4}',text, re.IGNORECASE)
            if reg is None:
                infoDict['DATE OF BIRTH'] = 'NA'
            else:
                d = reg[0].replace('-', ' ')
                d = d.split(':')
                #infoDict['DATE OF BIRTH'] = d
                if len(d) == 2:
                    d = d[1].strip()
                    infoDict['DATE OF BIRTH'] = d.replace(',','')
                else:
                    if d[0] == '':
                        infoDict['DATE OF BIRTH'] = 'NA'
                    else:
                        infoDict['DATE OF BIRTH'] = d[0].replace(',','')
        elif reg is None:
            infoDict['DATE OF BIRTH'] = 'NA'
        else:
            d = reg[0].replace('-', ' ')
            d = d.split(':')
            if len(d) == 2:
                d = d[1].strip()
                infoDict['DATE OF BIRTH'] = d.replace(',',' ')
            else:
                if d[0] == '':
                    infoDict['DATE OF BIRTH'] = 'NA'
                else:
                    d = reg[0].strip()
                    infoDict['DATE OF BIRTH'] = d.replace(',','')

    #Function to get Gender.
    def getGender(self, inputString, infoDict, file):
        reg = re.search(r'((((gender)|(sex))[\s]*:\s*((male)|(female)|(m)|(f)))|((male)|(female)))', inputString, re.IGNORECASE)
        ext = file[-4:]
        if reg is None and ext == '.pdf':
            text = convertPDFToTextUsingPypdf2(file)  # checking for text using PyPDF2 if not found using pdfminer
            text = text.replace('\n', '')
            reg = re.search(r'((((gender)|(sex))[\s]*:\s*((male)|(female)|(m)|(f)))|((male)|(female)))',text, re.IGNORECASE)
            if reg is None:
                infoDict['GENDER'] = 'NA'
            else:
                g = reg[0].replace('-', '')
                g = g.split(':')
                infoDict['GENDER'] = g
                if len(g) == 2:
                    infoDict['GENDER'] = g[1].strip()
                else:
                    if g[0] == '':
                        infoDict['GENDER'] = 'NA'
                    else:
                        infoDict['GENDER'] = g[0]
        elif reg is None:
            infoDict['GENDER'] = 'NA'
        else:
            g = reg[0].replace('-', '')
            g = g.split(':')
            if len(g) == 2:
                infoDict['GENDER'] = g[1].strip()
            else:
                if g[0] == '':
                    infoDict['GENDER'] = 'NA'
                else:
                    infoDict['GENDER'] = g[0]

    #Function to get Nationality.
    def getNationality(self, inputString, infoDict, file):
        reg = re.search(r'((nationality)|(country))[\s]*:\s*[a-zA-Z]*', inputString, re.IGNORECASE)
        ext = file[-4:]
        if reg is None and ext == '.pdf':
            text = convertPDFToTextUsingPypdf2(file)  # checking for text using PyPDF2 if not found using pdfminer
            text = text.replace('\n', '')
            reg = re.search(r'((nationality)|(country))[\s]*:\s*[a-zA-Z]*',text, re.IGNORECASE)
            if reg is None:
                infoDict['NATIONALITY'] = 'NA'
            else:
                nl = reg[0].replace('-', '')
                nl = nl.split(':')
                infoDict['NATIONALITY'] = nl
                if len(nl) == 2:
                    infoDict['NATIONALITY'] = nl[1].strip()
                else:
                    if nl[0] == '':
                        infoDict['NATIONALITY'] = 'NA'
                    else:
                        infoDict['NATIONALITY'] = nl[0]
        elif reg is None:
            infoDict['NATIONALITY'] = 'NA'
        else:
            nl = reg[0].replace('-', '')
            nl = nl.split(':')
            if len(nl) == 2:
                infoDict['NATIONALITY'] = nl[1].strip()
            else:
                if nl[0] == '':
                    infoDict['NATIONALITY'] = 'NA'
                else:
                    infoDict['NATIONALITY'] = nl[0]

    #Function to get Current Address
    def getCurrentAddress(self, inputString, infoDict, file):
        reg = re.search(r'(((place)|(current location)|(current address))\s*:\s[A-Za-z]*)', inputString,re.IGNORECASE)
        ext = file[-4:]
        if reg is None and ext == '.pdf':
            text = convertPDFToTextUsingPypdf2(file) # checking for text using PyPDF2 if not found using pdfminer
            text = text.replace('\n', '')
            reg = re.search(r'(((place)|(current location)|(current address))\s*:\s[A-Za-z]*)',text, re.IGNORECASE)
            if reg is None:
                infoDict['CURRENT ADDRESS'] = 'NA'
            else:
                adr = reg[0].replace('-', '')
                adr = adr.split(':')
                infoDict['CURRENT ADDRESS'] = adr
                if len(adr) == 2:
                    infoDict['CURRENT ADDRESS'] = adr[1].strip()
                else:
                    if adr[0] == '':
                        infoDict['CURRENT ADDRESS'] = 'NA'
                    else:
                        infoDict['CURRENT ADDRESS'] = adr[0]
        elif reg is None:
            infoDict['CURRENT ADDRESS'] = 'NA'
        else:
            adr = reg[0].replace('-', '')
            adr = adr.split(':')
            if len(adr) == 2:
                infoDict['CURRENT ADDRESS'] = adr[1].strip()
            else:
                if adr[0] == '':
                    infoDict['CURRENT ADDRESS'] = 'NA'
                else:
                    infoDict['CURRENT ADDRESS'] = adr[0]

if __name__ == "__main__":
    p = Parse()