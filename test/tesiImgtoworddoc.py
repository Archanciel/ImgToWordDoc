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
import docx

class TestImgToWordDoc(unittest.TestCase):
    def testSortFileNames(self):
        listOfFileNames = ['2.jpg', '1.png', '32.jpg', '8.png', 'name12.png']

        listOfFileNames.sort(key = imgToWordDoc.sortFileNames)
        self.assertEqual(['1.png', '2.jpg', '8.png', 'name12.png', '32.jpg'], listOfFileNames)


    def testSortFileNamesInvalidFileName(self):
        listOfFileNames = ['2.jpg', '1.png', '32.jpg', '8.png', 'name.png']

        with self.assertRaises(NameError):
            listOfFileNames.sort(key = imgToWordDoc.sortFileNames)

    def testCreateWordDocWithImgInDir(self):
        imgToWordDoc.createWordDocWithImgInDir()

        filesInDirList = os.listdir(currentdir)

        self.assertIn('test.docx', filesInDirList)

        os.remove(currentdir + '\\test.docx')


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
        doc = docx.Document(wordFilePathName)
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


    def testGetFilesInDir(self):
        curDir = os.getcwd()

        filesInDir = imgToWordDoc.getFilesInDir(curDir)
        print(filesInDir)