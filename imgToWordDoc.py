import os
import os.path as pa
from docx import Document
from docx.shared import Inches
from docx.shared import Cm
from PIL import Image
import re

IMG_MAX_WIDTH = 17.5    #anciennement 19.5
LATERAL_MARGIN = 2          #anciennement 1
SCREEN_DPI = 144    #on my 1920 x 1080' monitor


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
	curDir = os.getcwd()

	fileLst = os.listdir(curDir)
	imgFileLst = list(filter(lambda name: ".jpg" in name,fileLst))
	imgFileLst.sort(key=sortFileNames)

	doc = Document()

	#naming the created word file using the containing dir name
	targetWordFileName = curDir.split('\\')[-1]
	targetWordFileExt = ".docx"
	targetWordFileName = determineUniqueFileName(targetWordFileName, targetWordFileExt)

	setDocMargins(doc)
	i = 0

	for file in imgFileLst:
		#ajout d'un titre avant l'image
		doc.add_heading(str(i + 1) + '. A', level=1)

		#ajout de l'image. Si l'image est plus large que la largeur maximale, elle est réduite
		im = Image.open(file)
		imgWidthPixel, height = im.size
		imgWidthCm = imgWidthPixel / SCREEN_DPI * 2.54
		doc.add_picture(file, width=Cm(min(IMG_MAX_WIDTH, imgWidthCm)))

		#ajout d'un paragraphe bullet points
		paragraph = doc.add_paragraph('A')
		paragraph.style = 'List Bullet'
		i += 1

	fullTargetFileName = targetWordFileName + targetWordFileExt

	doc.save(fullTargetFileName)
	print("{0} file created with {1} image(s)".format(fullTargetFileName,i))


def sortFileNames(fileName):
	'''
	Using this function, a list of file names containing 1.jpg, 11.jpg, 2.jpg will 
	be ordered so: 1.jpg, 2.jpg, 11.jpg !
	:param fileName: 
	:return: number in img file name as int
	'''
	m = re.search(r'^(\d+).*', fileName)

	if m == None:
		raise NameError("Invalid img file name encountered: {0}. Img file names must start with a number for them to be inserted in the right order !".format(fileName))

	return int(m.group(1))


def setDocMargins(doc):
	sections = doc.sections

	for section in sections:
		section.top_margin = Cm(1)
		section.bottom_margin = Cm(1)
		section.left_margin = Cm(LATERAL_MARGIN)
		section.right_margin = Cm(LATERAL_MARGIN)


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

'''
Improvements:
° accept command line parm (use argparse)
° ajout images à la fin ou au début du doc ou à un index arbitraire: 0 == début, -1 == fin (default), ou n. 
  Dans tous les cas, on crée toujours un nouveau doc !
'''
