import argparse

parser = argparse.ArgumentParser(description="Add all images contained in current dir to a Word document. Each image "\
                                             "is added in a new paragraph. To facilitate further edition, the image "\
                                             "is preceded by a text line and followed by a bullet point section. "\
                                             "The images are added according to the alphabetic order of their "\
                                             "file names, so use names starting by a number (i.e. 1.jpg, 2.jpg, ...). "\
                                             "If no document name is specified, the created document has "\
                                             "the same name as the containing dir. An existing document with "\
                                             "same name is never overwritten. Instead, a new document with a "\
                                             "name incremented by 1 (i.e. myDoc1.docx, myDoc2.docx, ...) "\
                                             "is created.")
parser.add_argument("-d", "--document", nargs="?", help="existing document to which the current dir images are "\
                                                   "to be added. For your convinience, the initial document is "\
                                                   "not modified. Instead, the original document is copied with a "\
                                                   "name incremented by one and the images are added to the copy.")
parser.add_argument("-i", "--insertionPos", type=int, nargs="?", default=-1, help="paragraph number BEFORE which to insert the "\
                                                                          "images. default value is -1 --> end of document. "\
                                                                          "1 --> start of document (before paragraph 1). "\
                                                                          "The insertion position is ignored if no existing "\
                                                                          "document is specified !")
parser.add_argument("-p", "--pictures", nargs="*", help="paragraph number BEFORE which to insert the "\
                                                                          "images. default value is -1 --> end of document. "\
                                                                          "1 --> start of document (before paragraph 1). "\
                                                                          "The insertion position is ignored if no existing "\
                                                                          "document is specified !")
args = parser.parse_args()

print(args.document)
print(args.insertionPos)
print(args.pictures)