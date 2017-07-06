import os
import os.path as pa
from docx import Document
from docx.shared import Inches
from docx.shared import Cm

IMG_WIDTH = 19.5
MARGIN = 1
SCREEN_DPI = 144


def createWordDocWithImgInDir():
	curDir = os.getcwd()

	fileLst = os.listdir(curDir)
	imgFileLst = list(filter(lambda name: ".jpg" in name,fileLst))

	doc = Document()

	targetWordFileName = "hello"
	targetWordFileExt = ".docx"
	targetWordFileName = determineUniqueFileName(targetWordFileName, targetWordFileExt)

	setDocMargins(doc)

	for file in imgFileLst:
		imgPixels = 12
		imgWidthCm = imgPixels / SCREEN_DPI * 2.54
		doc.add_picture(file, width=Cm(IMG_WIDTH))
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