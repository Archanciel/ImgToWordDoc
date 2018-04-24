from docx import Document
from docx.shared import Inches

document = Document()

p = document.add_paragraph()
r = p.add_run()
r.add_text('Good Morning every body,This is my ')
picPath = 'D:/Development/Python/aa.png'
r.add_picture(picPath)
r.add_text(' do you like it?')

document.save('demo.docx')

'''
Here's my solution. It has the advantage on the first proposition that it surrounds 
the picture with a title (with style Header 1) and a section for additional comments. 
Note that you have to do the insertions in the reverse order they appear in the Word 
document.

This snippet is particularly useful if you want to programmatically insert pictures 
in an existing document.
'''

document = Document()

p = document.add_paragraph('Picture bullet section', 'List Bullet')
p = p.insert_paragraph_before('')
r = p.add_run()
r.add_picture(picPath)
p = p.insert_paragraph_before('My picture title', 'Heading 1')

document.save('demo_better.docx')