#THis is a Python3 script to give a word choice check of documents.
#Every word in the (*.txt or *.doc) file is put into a list of tuples (String word, int count)
#Every word in a paragraph is put into a list of tuples (String word, int count)
#Every word in a sentence is put into a list of tuples (String word, int count)
#Every word is then scored accordingly in a dictionary:
#     for every word a sentence contains more than once, it receives a score of +3
#     for every word a paragraph contains more than once, it recieves a score of +2
#     for every word that is used in a file, it receive a score of its total word count.
#The user is then presented these words in order of score, and is allowed to cycle back and forth through the list of words
#The user may choose to exit the application by entering "0" at any time.

############################################################
####################    IMPORTS         ####################
############################################################
import os
from docx import Document

############################################################
####################    FUNCTIONS       ####################
############################################################
def getFileAttributes():
    #fileName = input("First, enter the text you wish to check into a text file (*.txt).\nThen, please enter the file extension and file name (ex C:\\Users\\Tom\\Desktop\\Some File.txt ):\n\t")
    #while (True):
    filePath = os.path.dirname(os.path.realpath(__file__))
    fullFileName = input("Enter the name and extension of the file (ex. Some File.docx).\nIMPORTANT:  It must be in the same directory as this program!:\n\n    >>>" + filePath + "\\")
    fileName, fileExtension = "", ""
    beforeExtension = True
    if (fullFileName.count(".") != 1):
        print("ERROR, PROBLEM WITH FILE EXTENSION (check period!).\n\n\n")
        return getFileAttributes()
    for char in fullFileName:
        if char != "." and beforeExtension:
            fileName += char
        else:
            beforeExtension = False
            fileExtension += char
    #fileString = filePath + "\\" + fileName
    if fileExists(filePath, fileName, fileExtension) == False:
        print("ERROR, THAT FILE COULD NOT BE FOUND.\n\n\n")
        return getFileAttributes()
    return (filePath, fileName, fileExtension)#return fileString#return f


def fileExists(filePath, fileName, fileExtension, action = 'r'):
    try:
        f = open(filePath + "\\" + fileName + fileExtension, action)
        f.close()
        return True;
    #except FileNotFoundError:
    except:
        try:
            doc = Document(open(filePath + "\\" + fileName + fileExtension, action + 'b'))
            doc.close()
            return True
        except FileNotFoundError:
            return False
        return False
    return True


def addNewLine(s, fileName, allowDuplicates = True, alphabetizeFile = False):#return none;  appends string s to a new line in file fileName and then rewrites the file so it is in alpabetical order.  CURRENTLY NOT EFFIECIENT FOR LARGE FILE SIZES.  Assumes that s is a verified string.
    if isKnown(s, fileName):#checks if the element to be added is already in the file.  Also checks that there is a file to add to!  (if not, creates the file!)
        #f = open(fileName, 'r')#read this file
        if allowDuplicates == False:#from encasing (previous) if statement, know that there is a dupliate in there! So no need to add, just return!
            return#no need to add anthing
        f = open(fileName, 'a')#append this file
        f.write(s + "\n")#add the element to the file before sorting it.
        f.close()
    else: #file (or string in that file) does not exist!  no need to D2!
        f = open(fileName, 'a')#append this file
        f.write(s + "\n")#add the element to the file before sorting it.
        print("Appended " + s + " to " + fileName)
        f.close()
    if alphabetizeFile == True:
        #need to enter the name so it is in alphabetical order!
        f = open(fileName, 'rb')
        tmpList = []
        for line in sorted(f):
            tmpList.append(line)
            print("#####" + line)
        f.close()
        f = open(fileName, 'w')
        for element in tmpList:
            f.write(element)
        f.close()


def createFile(fileName):
    f = open(fileName, 'a')
    f.close


def fileToString(directory, fileName, extension):
    def isAscii(s):
        try:
            s.encode('ascii')
        except UnicodeEncodeError:
            return False
        else:
            return True
    #def isAscii(s):
    #    import string
    #    for c in s:
    #        if c not in string.ascii_letters and c != "\n" and c != "\r\n" and c != " ":
    #            print("NOT ASCII or return!")
    #            return False
    #        else:
    #            print(c)
    #    return True
    ############
    ### MAIN ###
    ############
    eraseTmpFile = False
    TMP_FILENAME = directory + "tmpFile" + '.txt'
    if (extension == ".txt"):
        f = open(directory + "\\" + fileName + extension, 'r')
    elif (extension == ".docx"):
        print ("\n Looking for the following doc: (" + directory + "\\) " + fileName + extension + "\n")
        doc = Document(open(directory + "\\" + fileName + extension, 'rb'))
        createFile(TMP_FILENAME)
        eraseTmpFile = True
        tmpf = open(TMP_FILENAME, 'a')
        for para in doc.paragraphs:
            #for run in para.runs:#get string from doc.
            #    tmpf.write(run.text)#turn string to txt file
            try:
                tmpf.write(para.text)
            except UnicodeEncodeError:
                tmpf.write(" ")
            tmpf.write("\n")#turn string to txt file
        tmpf.close()
        f = open(TMP_FILENAME, 'r')
        #eraseTmpFile = True
    s = "";
    for line in f:
        for char in line:
            #isascii = lambda s: len(tmp_s) == len(tmp_s.encode())
            #print(char)
            if (isAscii(str(char))):
                s += str(char)
            else:
                s+= " "
    f.close()
    if (eraseTmpFile):
        os.remove(TMP_FILENAME)
    return s

def updateScores(l, d, bonus):#can this be pass by ref? also may be good idea to use the count() method of lists!
    alreadyUsedElts = [" "]
    alreadyUsedElts.clear()
    for elt in l:
        if elt in d:
            if alreadyUsedElts.count(elt) > 0:
                d[elt] += bonus
            else:
                alreadyUsedElts.append(elt)
                d[elt] += 1
            #print(d)
            #input(elt + " " + str(d[elt]) + ".  Bonus = " + str(bonus))
        else:
            alreadyUsedElts.append(elt)
            d[elt] = 1
            #input(elt + " " + str(d[elt]) + ".  Bonus = " + str(bonus))
            #print(d)
            #input()
    #print(sorted(d), "\n\n", d)
    #input()
    return d


def delimitStringToList(s):
    FIRST_ELT = " "
    l = [FIRST_ELT]
    l.clear()
    tmpStr = ""
    for char in s:
        #print ("char" + char + ";    tmpStr = " + tmpStr)
        if isRunonPunctuation(char) or isMiscPunctuation(char) or char == " ":
            if tmpStr != "":
                l.append(tmpStr)
            #print("1appending: \'" + tmpStr + "\'")
            #print("+__________________")
            tmpStr = ""
        elif char == "\n":
            if tmpStr != "":
                l.append(tmpStr)
            #print("2appending: \'" + tmpStr + "\'")
            l.append(char)
            #print("2appending: \'" + char + "\'")
            #print("+__________________")
            tmpStr = ""
        else:
            tmpStr += char
    if tmpStr != "":
        l.append(tmpStr)
    #print("3appending: \'" + tmpStr + "\'")
    #print("+__________________")
    return l

def isEndingPunctuation(c):
    return c == '.' or c == '!' or c == '?' or c == ';'
def isRunonPunctuation(c):
    return c == "," or c == "\t" or c == ':'
def isMiscPunctuation(c):
    return c == "\"" or c == "(" or c == ")"# or c == "\'"
def isPunctuation(c):
    return isEndingPunctuation(c) or isRunonPunctuation(c) or isMiscPunctuation(c)

def convertToSortedList(d):
    l = [""]
    l.clear()
    for k,v in d.items():
        l.append((v, k))
    return sorted(l, key=lambda x: x[0])

def getLength(l):
    totalWords = 0
    for elt in l:
        totalWords += 1
    return totalWords


############################################################
####################        MAIN        ####################
############################################################
from docx import Document
FIRST_ELT = " "
SENTENCE_WEIGHT = 10;
PARAGRAPH_WEIGHT = 4;
FILE_WEIGHT = 1;
f = getFileAttributes()
while(True):
    s = fileToString(f[0], f[1], f[2]).upper()
    allWords = delimitStringToList(s)#takes out all spaces, only leaving returns
    ###THere has to be a better way to initialize empty lists...
    fileWordCount = [FIRST_ELT]
    fileWordCount.clear()
    paragraphWordCount = [FIRST_ELT]
    paragraphWordCount.clear()
    sentenceWordCount = [FIRST_ELT]
    sentenceWordCount.clear()

    allWordScores = {}
    for word in allWords:
        #remove punctuation from the end of words
        if isEndingPunctuation(word[-1]) == False:
            sentenceWordCount.append(word)
        else:#move onto new sentence
            word = word[:-1]
            sentenceWordCount.append(word)#add the last word minus the punctuation
            allWordScores = updateScores(sentenceWordCount, allWordScores, SENTENCE_WEIGHT)
            sentenceWordCount.clear()

        #separate into paragraphs (the returns have already been removed and are treated as individual words!)
        if word != "\n" and word != "\r\n":
            paragraphWordCount.append(word)
        else:#move onto new paragraph
            allWordScores = updateScores(paragraphWordCount, allWordScores, PARAGRAPH_WEIGHT)
            paragraphWordCount.clear()

        #update file Count
        fileWordCount.append(word)

    allWordScores = updateScores(fileWordCount, allWordScores, FILE_WEIGHT)
    allWordScores = updateScores(paragraphWordCount, allWordScores, PARAGRAPH_WEIGHT)
    allWordScores = updateScores(sentenceWordCount, allWordScores, SENTENCE_WEIGHT)

    allWords = {}
    allWords = updateScores(fileWordCount, allWords, 1)

    print("\n+____________________________________________\n")
    print (" _Freq_\t_Total_\t_Word_")
    totalWords = getLength(convertToSortedList(allWordScores))
    count = 0
    for elt in convertToSortedList(allWordScores):
        if elt[1] != "\n":
            print (" " + str(elt[0]) + "\t" + str(allWords[elt[1]]) + "\t" + elt[1])
        else:
            print (" " + str(elt[0]) + "\t" + str(allWords[elt[1]]) + "\t" + "<CARRIAGE RETURN>")
        count += 1
        if ((totalWords - count) % 25 == 0) and (totalWords - count >= 5):
            print(" ----------- Top " + str(totalWords - count) + " -----------")

    print("Scoring complete.\nTotal Vocab used: " + str(totalWords) + " different words.\n")
    input("Hit \'enter\' to refresh the evaluation\n\n")
print("Goodbye.")
exit(0)
