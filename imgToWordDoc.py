import os
import os.path as pa
from docx import Document
from docx.shared import Inches
from docx.shared import Cm
from PIL import Image

IMG_WIDTH = 19.5
MARGIN = 1
SCREEN_DPI = 144    #on my 1920 x 1080' monitor


def createWordDocWithImgInDir():
	'''
	Python utility to add all the images of a directory to a new Word document in order to facilitate
	documentation creation. The images are added in their file name ascending order.
	'''
	curDir = os.getcwd()

	fileLst = os.listdir(curDir)
	imgFileLst = list(filter(lambda name: ".jpg" in name,fileLst))
	imgFileLst.sort()

	doc = Document()

	targetWordFileName = "hello"
	targetWordFileExt = ".docx"
	targetWordFileName = determineUniqueFileName(targetWordFileName, targetWordFileExt)

	setDocMargins(doc)

	for file in imgFileLst:
		im = Image.open(file)
		imgWidthPixel, height = im.size
		imgWidthCm = imgWidthPixel / SCREEN_DPI * 2.54
		doc.add_picture(file, width=Cm(min(IMG_WIDTH, imgWidthCm)))
		doc.add_paragraph()

	doc.save(targetWordFileName + targetWordFileExt)


def setDocMargins(doc):
	sections = doc.sections

	for section in sections:
		section.top_margin = Cm(MARGIN)
		section.bottom_margin = Cm(MARGIN)
		section.left_margin = Cm(MARGIN)
		section.right_margin = Cm(MARGIN)


def determineUniqueFileName(targetWordFileName, targetWordFileExt):
	'''Verify if a file with same name exists and increment the name by one in this case.

	Ex: if hello.docx exists, returns hello1, hello2, etc
	'''
	i = 1
	lookupWordFileName = targetWordFileName

	while pa.isfile(lookupWordFileName + targetWordFileExt):
		lookupWordFileName = targetWordFileName + str(i)
		i += 1

	return lookupWordFileName


if __name__ == '__main__':
	createWordDocWithImgInDir()