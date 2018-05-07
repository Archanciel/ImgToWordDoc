import os
import re

import os.path as curDir
from PIL import Image
from docx import Document
from docx.shared import Cm
import argparse

IMG_MAX_WIDTH = 17.5  # anciennement 19.5
LATERAL_MARGIN = 2  # anciennement 1
SCREEN_DPI = 144  # on my 1920 x 1080' monitor
WORD_FILE_EXT = ".docx"


def getCommandLineArgs():
    '''
    Uses argparse to acquire the user optional command line arguments.

    :return: document name (may be None) and insertion point
    '''
    parser = argparse.ArgumentParser(
        description="Add all images contained in current dir to a Word document. Each image " \
                    "is added in a new paragraph. To facilitate further edition, the image " \
                    "is preceded by a text line and followed by a bullet point section. " \
                    "The images are added according to the alphabetic order of their " \
                    "file names, so use names starting by a number (i.e. 1.jpg, 2.png, ...). " \
                    "If no document name is specified, the created document has " \
                    "the same name as the containing dir. An existing document with " \
                    "same name is never overwritten. Instead, a new document with a " \
                    "name incremented by 1 (i.e. myDoc1.docx, myDoc2.docx, ...) " \
                    "is created.")
    parser.add_argument("-d", "--document", nargs="?", help="existing document to which the current dir images are " \
                                                            "to be added. For your convinience, the initial document is " \
                                                            "not modified. Instead, the original document is copied with a " \
                                                            "name incremented by one and the images are added to the copy.")
    parser.add_argument("-i", "--insertionPos", type=int, nargs="?", default=-1,
                        help="paragraph number BEFORE which to insert the " \
                             "images. default value is -1 --> end of document. " \
                             "1 --> start of document (before paragraph 1). ")
    args = parser.parse_args()

    return args.document, args.insertionPos


def openExistingOrCreateNewWordDoc(documentName):
    '''
    Opens the passed Word doc documentName located in the current dir. If no document with the passed name
    exist in the current dir, a new empty Word document is created.
    by 1.
    :param userDocumentName:

    :return: either existing or brand new document.
    '''
    if curDir.isfile(documentName + WORD_FILE_EXT):
        return Document(documentName)
    else:
        return Document()


def getFilesInDir(directory):
    '''
    Returns the list of file names contained in the passed directory
    :param directory:
    :return:
    '''
    fileList = []

    for fname in os.listdir(directory):
        path = os.path.join(directory, fname)
        if os.path.isdir(path):
            # skip directories
            continue
        else:
            fileList.append(fname)

    return fileList

def createWordDocWithImgInDir():
    '''
    Python utility to add all the images of a directory to a new Word document in order to facilitate
    documentation creation. The images are added in their file name ascending order.

    *** USAGE ***

    In a command window opened on the dir containing the images, after copying the imgToWordDoc.py file
    in it, simply type

    python imgToWordDoc.py

    This will create a new Word document whose name is the name of the current dir. In case
    the dir already contains a Word documant with the same name, an incremented number is
    appended to the file name !
    '''
    userDocumentName, userInsertionPos = getCommandLineArgs()

    curDir = os.getcwd()

    fileLst = getFilesInDir(curDir)
    imgFileLst = list(filter(lambda name: ".jpg" in name or ".png" in name, fileLst))
    imgFileLst.sort(key=sortFileNames)
    doc = None
     
    if userDocumentName:
        #user provided a document name. Either open the existing document
        #or create a new empty one
        if WORD_FILE_EXT in userDocumentName:
            targetWordFileName = userDocumentName[:-5]
        else:
            targetWordFileName = userDocumentName
        doc = openExistingOrCreateNewWordDoc(targetWordFileName)
    else:
        #no document name provided, so name the created word file 
        #using the containing dir name
        if os.name == 'posix':
            targetWordFileName = curDir.split('/')[-1]
        else:
            targetWordFileName = curDir.split('\\')[-1]
        doc = Document()
        
    targetWordFileName = determineUniqueFileName(targetWordFileName)
 
    setDocMargins(doc)
    i = 0

    for fileName in imgFileLst:
        # ajout d'un titre avant l'image
        doc.add_heading('A', level=1)

        # ajout de l'image. Si l'image est plus large que la largeur maximale, elle est réduite
        im = Image.open(fileName)
        imgWidthPixel, height = im.size
        imgWidthCm = imgWidthPixel / SCREEN_DPI * 2.54
        doc.add_picture(fileName, width=Cm(min(IMG_MAX_WIDTH, imgWidthCm)))

        # ajout d'un paragraphe bullet points
        paragraph = doc.add_paragraph('A')
        paragraph.style = 'List Bullet'
        i += 1

    doc.save(targetWordFileName)
    resultMsg = "{0} file created with {1} image(s). Manually add auto numbering to the 'Header 1' / 'Titre 1' style !".format(
        targetWordFileName, i)
    print(resultMsg)

    return resultMsg
 

def sortFileNames(fileName):
    '''
    Using this function, a list of file names containing 1.jpg, 11.jpg, 2.jpg will
    be ordered so: 1.jpg, 2.jpg, 11.jpg !
    :param fileName:
    :return: number in img file name as int
    '''
    m = re.search(r'(\d+).*', fileName)

    if m == None:
        raise NameError(
            "Invalid img file name encountered: {0}. Img file names must contain a number for them to be inserted in the right order (depends on this number) !".format(
                fileName))

    return int(m.group(1))


def setDocMargins(doc):
    sections = doc.sections

    for section in sections:
        section.top_margin = Cm(1)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(LATERAL_MARGIN)
        section.right_margin = Cm(LATERAL_MARGIN)


def determineUniqueFileName(wordFileName):
    '''Verify if a file with same name exists and increment the name by one in this case.

    Ex: if hello.docx exists, returns hello1, hello2, etc

    :param  wordFileName without extention
    :return wordFileName + incremented number (if wordFileName exists in curr dir) +
            word file extention
    '''
    i = 1
    lookupWordFileName = wordFileName

    while curDir.isfile(lookupWordFileName + WORD_FILE_EXT):
        lookupWordFileName = wordFileName + str(i)
        i += 1

    return lookupWordFileName + WORD_FILE_EXT


if __name__ == '__main__':
    try:
        createWordDocWithImgInDir()
    except NameError as e:
        print(e)
