#Import for parsing pdf
from pdfminer.pdfparser import PDFParser, PDFDocument
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextBox, LTTextLine
#Import for parsing doc
import docx2txt

class Parser:
    def __init__(self, fileName):
        #Grab from string type of file
        typeFile = fileName.split('.')[-1]
        #Choosing method
        if typeFile == 'pdf' or typeFile == 'PDF':
            self.allText = self.pdfParse(fileName)
        if typeFile == 'docx' or typeFile == 'DOCX':
            self.allText = self.docxParse(fileName)
        if typeFile == 'txt' or typeFile == 'TXT':
            self.allText = self.txtParse(fileName)

    def txtParse(self, txtName):
        #Open file
        fileTxt = open(txtName)
        matrixDoc = []
        try:
            for line in fileTxt:
                #Deleting trash tabs and empty strings (\n)
                if line == '\n':
                    continue
                line = line.strip('\t')
                line = line.strip('\n')
                matrixDoc.append(line)
        finally:
            #Close file
            fileTxt.close()
        #Cleaning from empty strings
        matrixDoc = list(filter(bool, matrixDoc))
        for i in range(len(matrixDoc)):
            matrixDoc[i] = matrixDoc[i].split(' ')
        return matrixDoc

    def docxParse(self, docName):
        #Grab text from docx
        matrixDoc= docx2txt.process(docName)
        #Split to lines, then to words
        matrixDoc = matrixDoc.split('\n')
        #Cleaning from empty strings
        matrixDoc = list(filter(bool, matrixDoc))
        #Spliting on words
        for i in range(len(matrixDoc)):
            matrixDoc[i] = matrixDoc[i].split(' ')
        return matrixDoc

    def pdfParse(self, pdfName):
        #Open file in binary format
        fileInBinary = open(pdfName, 'rb')
        #Create a PDF parser object associated with the file
        parser = PDFParser(fileInBinary)
        # Create a PDF document object that stores the document structure.
        document = PDFDocument()
        # Connect the parser and document objects.
        parser.set_document(document)
        document.set_parser(parser)
        #Password for pdf
        document.initialize('')
        #Try parse document, is it allows text extraction
        if not document.is_extractable:
            raise PDFTextExtractionNotAllowed
        # Create a PDF resource manager object that stores shared resources.
        resourceManager = PDFResourceManager()
        #Parameters for spacing
        laparams = LAParams()
        # Create a PDF device object.
        device = PDFPageAggregator(resourceManager, laparams=laparams)
        # Create a PDF interpreter object.
        interpreter = PDFPageInterpreter(resourceManager, device)
        # Process each page contained in the document.
        matrixDoc = ''
        for page in document.get_pages():
            interpreter.process_page(page)
            layout = device.get_result()
            for lt_obj in layout:
                #LTTextBox it is box with text, LTTextLine ->line of text
                if isinstance(lt_obj, LTTextBox) or isinstance(lt_obj, LTTextLine):
                    matrixDoc += lt_obj.get_text()
        #Create matrix from one peace of text
        matrixDoc = matrixDoc.split('\n')
        #Cleaning from empty strings
        matrixDoc = list(filter(bool, matrixDoc))
        for i in range(len(matrixDoc)):
            matrixDoc[i] = matrixDoc[i].split(' ')
        #Close file
        fileInBinary.close()
        return matrixDoc

    def search(self, keyWord):
        linesWithKeyWords = []
        #Put keyWord symbols in lowercase
        keyWord = keyWord.lower()
        #List of badSymbols , delete from words
        badSymbols = "!@.,/#$%:;'?()-"
        #Get lines from allText
        for i in range(len(self.allText)):
            #print(self.allText[i])
            #Cleaning words from badSymbols, and put in lowercase
            for word in self.allText[i]:
                for char in badSymbols:
                    word = word.replace(char , '')
                    word = word.lower()
                #Append 5 lines in result list
                if keyWord == word:
                    for j in [-2,-1,0,1,2]:
                        linesWithKeyWords.append(self.allText[i+j])
                    #TODO Delete this
                    linesWithKeyWords.append('\n')
        return linesWithKeyWords


obj=Parser('jad.docx')
for line in obj.search('повітря'):
    print (line)
obj2=Parser('jad.txt')
for line in obj2.search('Осінь'):
    print (line)
obj3=Parser('jad.pdf')
for line in obj3.search('винна'):
    print (line)
