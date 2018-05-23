# ImgToWordDoc
Python command line utility used to add all images of a directory to a new or existing Word document in order to facilitate the creation of documentation.
The images are added according to the ascending order of the number contained anywhere in their file name.

## Detailed usage help
### python ImgToWordDoc.py [-h] [-d [DOCUMENT]] [-i [INSERTIONPOS]] [-p [PICTURES [PICTURES ...]]]

Adds or inserts all or part of the images contained in the current dir to a
Word document. Each image is added in a new paragraph. To facilitate further
edition, the image is preceded by a header line and followed by a bullet point
section. The images are added according to the ascending order of the number
contained in their file name. An error will occur if one of the image file
name does not contain a number (valid image file names are: 1.jpg, image2.jpg,
3.png, ...). If no document name is specified, the created document has the
same name as the containing dir. An existing document with same name is never
overwritten. Instead, a new document with a name incremented by 1 (i.e.
myDoc1.docx, myDoc2.docx, ...) is created. Using the utility in add mode, i.e.
without specifying an insertion point, creates a new document in which the
specified images will be added. If the current dir already contains a document
with images and comments you want to keep, use the insertion parameter which
will insert the new images at the specified position and preserve the initial
content. Without using the -p parameter, all the images of the current dir are
collected for the addition/insertion. -p enables to specify precisely the
images to add/insert using only the numbers contained in the image file names.

#### optional arguments:
##### -h, --help
              show this help message and exit
##### -d [DOCUMENT], --document [DOCUMENT]
              existing document to which the images are to be added.
              For your convenience, the initial document is not
              modified. Instead, the original document is copied
              with a name incremented by one and the images are
              added/inserted to the copy.
##### -i [INSERTIONPOS], --insertionPos [INSERTIONPOS]
              paragraph number BEFORE which to insert the images. 1
              --> start of document (before paragraph 1). 0 --> end
              of document.
##### -p [PICTURES [PICTURES ...]], --pictures [PICTURES [PICTURES ...]]
              numbers contained in the image file names which are
              selected to be inserted in the existing document.
              Exemple: -p 1 8 4-6 9-10 or -p 1,8, 4-6, 9-10 means
              the images whose name contain the specified numbers
              will be added or or inserted in ascending number
              order, in this case 1, 4, 5, 6, 8, 9, 10. If this parm
              is omitted, all the pictures in the curreent dir are
              added or inserted.
## Required libraries
* python-docx
* pillow
