import docx

def getText(filename):
    doc = docx.Document(filename)
    fullText = []
    i = 1

    for para in doc.paragraphs:
        if para.style.name == 'Heading 1':
            fullText.append(para._body)
            if i == 2:
                para.insert_paragraph_before('Lorem ipsum')
            i += 1
    doc.save(filename)
    return '\n'.join(fullText)

print(getText('ImgToWordDoc4.docx'))

