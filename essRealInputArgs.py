import argparse

parser = argparse.ArgumentParser(description="Add all images contained in current dir to a Word document. "\
                                             "If no document name is specified, the created document has "\
                                             "the same name as the containing dir. An existing document with "\
                                             "same name is never overwritten. Instead, a new document with a "\
                                             "name incremented by 1 is created.")
parser.add_argument("-d", "--document", nargs="?", help="existing document name to which the current dir images are "\
                                                   "to be added. If not specified, the newly created doc will only "\
                                                   "contain the current dir images.")
parser.add_argument("-i", "--insertionPos", type=int, nargs="?", default=-1, help="paragraph number before which to insert the "\
                                                                          "images. -1 (end) by default. 1 --> doc start. "\
                                                                          "The insertion pos is ignored if no existing "\
                                                                          "document is specified !")
args = parser.parse_args()

print(args.document)
print(args.insertionPos)