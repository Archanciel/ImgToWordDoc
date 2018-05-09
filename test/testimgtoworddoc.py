import inspect
import os
import sys
import unittest

currentdir = os.path.dirname(os.path.abspath(inspect.getfile(inspect.currentframe())))
parentdir = os.path.dirname(currentdir)
sys.path.insert(0, parentdir)
sys.path.insert(0,currentdir) # this instruction is necessary for successful importation of utilityfortest module when
                              # the test is executed standalone

import imgToWordDoc
from docx import Document

class TestImgToWordDoc(unittest.TestCase):
    def testSortNumberedStringsFunc(self):
        listOfFileNames = ['2.jpg', '1.png', '32.jpg', '8.png', 'name12.png']

        listOfFileNames.sort(key = imgToWordDoc.sortNumberedStringsFunc)
        self.assertEqual(['1.png', '2.jpg', '8.png', 'name12.png', '32.jpg'], listOfFileNames)


    def testSortNumberedStringsFuncInvalidFileName(self):
        listOfFileNames = ['2.jpg', '1.png', '32.jpg', '8.png', 'name.png']

        with self.assertRaises(NameError):
            listOfFileNames.sort(key = imgToWordDoc.sortNumberedStringsFunc)


    def testCreateWordDocWithImgInDirIncrementFileName(self):
        imgToWordDoc.createWordDocWithImgInDir()
        imgToWordDoc.createWordDocWithImgInDir()

        filesInDirList = os.listdir(currentdir)

        self.assertIn('test.docx', filesInDirList)
        self.assertIn('test1.docx', filesInDirList)

        os.remove(currentdir + '\\test.docx')
        os.remove(currentdir + '\\test1.docx')


    def testCreateWordDocWithImgInDirImgNumber(self):
        returnedInfo = imgToWordDoc.createWordDocWithImgInDir()

        self.assertEqual("test.docx file created with 3 image(s). Manually add auto numbering to the 'Header 1' / 'Titre 1' style !", returnedInfo)
        wordFilePathName = currentdir + '\\test.docx'
        doc = Document(wordFilePathName)
        self.assertEqual(9, len(doc.paragraphs)) # 3 headers + 3 images + 3 bullet points sections

        os.remove(wordFilePathName)


    def testDetermineUniqueFileNameNoWordFileExist(self):
        wordFileName = "notExistFileName"
        wordFileNameWithExt = imgToWordDoc.determineUniqueFileName(wordFileName)
        self.assertEqual("notExistFileName.docx", wordFileNameWithExt)


    def testDetermineUniqueFileNameWordFileExistNoSuffixNumber(self):

        # create the situation where there's 'existingFileName.docx' only
        # in the current dir

        wordFileName = "existingFileName"

        fileNameExt = wordFileName + ".docx"

        with open(fileNameExt, 'w') as f:
            pass

        # perform the test

        wordFileNameWithExt = imgToWordDoc.determineUniqueFileName(wordFileName)
        self.assertEqual("existingFileName1.docx", wordFileNameWithExt)

        # clean up

        os.remove(fileNameExt)


    def testDetermineUniqueFileNameTwoWordFileExistOneWithSuffixNumber(self):

        # create the situation where there's 'existingFileName.docx' and 'existingFileName1.docx'
        # in the current dir

        wordFileName = "existingFileName"
        fileNameExt = wordFileName + ".docx"

        with open(fileNameExt, 'w') as f:
            pass

        wordFileNameSuffixOne = wordFileName + '1'
        fileNameExtSuffixOne = wordFileNameSuffixOne + ".docx"

        with open(fileNameExtSuffixOne, 'w') as f:
            pass

        # perform the test

        wordFileNameWithExt = imgToWordDoc.determineUniqueFileName(wordFileName)
        self.assertEqual("existingFileName2.docx", wordFileNameWithExt)

        # clean up

        os.remove(fileNameExt)
        os.remove(fileNameExtSuffixOne)


    def testGetFilesInTestDir(self):
        curDir = os.getcwd()

        filesInDir = imgToWordDoc.getFilesInDir(curDir)
        self.assertEqual(6, len(filesInDir))


    def testGetFilesInEmptyDir(self):
        os.makedirs('empty')
        os.chdir('empty')
        curDir = os.getcwd()

        filesInDir = imgToWordDoc.getFilesInDir(curDir)
        self.assertEqual(0, len(filesInDir))

        os.chdir('..')
        os.removedirs('empty')


    def testGetSortedImageFileNames(self):
        curDir = os.getcwd()
        imgFileLst = imgToWordDoc.getSortedImageFileNames(curDir)
        self.assertEqual(['1.jpg', 'name3.jpg', '4.jpg'],imgFileLst)


    def testGetSortedImageFileNamesWithInvalidFileName(self):
        '''
        Current dir contains image file whose name contains no number.
        :return:
        '''
        curDir = os.getcwd()

        invalidFileName = 'invalFileName.jpg'

        with open(invalidFileName, 'w') as f:
            pass

        with self.assertRaises(NameError):
            imgToWordDoc.getSortedImageFileNames(curDir)

        os.remove(invalidFileName)


    def testDetermineInsertionPointInExistingWordDocWithOneParagraph(self):
        wordDoc = Document('oneImg.docx')
        insertionPoint = 1
        firstParagraph = wordDoc.paragraphs[0]

        titleString = firstParagraph.text
        self.assertEqual('My picture title', titleString)
        self.assertEqual(titleString, imgToWordDoc.determineInsertionPoint(insertionPoint, wordDoc).text)


    def testDetermineInsertionPointInExistingWordDocWithTwoParagraphsPos1(self):
        wordDoc = Document('twoImg.docx')
        insertionPoint = 1
        firstParagraph = wordDoc.paragraphs[0]

        titleString = firstParagraph.text
        self.assertEqual('My picture one title', titleString)
        self.assertEqual(titleString, imgToWordDoc.determineInsertionPoint(insertionPoint, wordDoc).text)


    def testDetermineInsertionPointInExistingWordDocWithTwoParagraphsPos2(self):
        wordDoc = Document('twoImg.docx')
        insertionPoint = 2
        firstParagraph = wordDoc.paragraphs[3] # 3 headers + 3 images + 3 bullet points sections

        titleString = firstParagraph.text
        self.assertEqual('My picture two title', titleString)
        self.assertEqual(titleString, imgToWordDoc.determineInsertionPoint(insertionPoint, wordDoc).text)
