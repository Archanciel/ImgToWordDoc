import os
import os.path as pa
from docx import Document
from docx.shared import Inches
from docx.shared import Cm
from PIL import Image
import re

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

	doc = Document()

	targetWordFileName = "hello"
	targetWordFileExt = ".docx"
	targetWordFileName = determineUniqueFileName(targetWordFileName, targetWordFileExt)

	setDocMargins(doc)
	i = 0

	for file in imgFileLst:
		im = Image.open(file)
		imgWidthPixel, height = im.size
		imgWidthCm = imgWidthPixel / SCREEN_DPI * 2.54
		doc.add_picture(file, width=Cm(min(IMG_WIDTH, imgWidthCm)))
		doc.add_paragraph()
		i += 1

	fullTargetFileName = targetWordFileName + targetWordFileExt

	doc.save(fullTargetFileName)
	print("{0} file created with {1} image(s)".format(fullTargetFileName,i))


	'''
	Using this function, a list of file names containing 1.jpg, 11.jpg, 2.jpg will 
	be ordered so: 1.jpg, 2.jpg, 11.jpg !
	:param fileName: 
	:return: 
	'''
	m = re.search(r'^(\d+).*', fileName)

	if m == None:
		raise NameError("Invalid img file name encountered: {0}. Img file names must start with a number for them to be inserted in the right order !".format(fileName))

	return int(m.group(1))


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
	try:
		createWordDocWithImgInDir()
	except NameError as e:
		print(e)
