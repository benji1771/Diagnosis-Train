from docx import Document
import re
import os

allDiagn = Document()

for files in os.listdir():
    if files.endswith('docx'):
        doc = Document(files)
        diagn = ''
        for num in range (3):
            diagn = doc.paragraphs[num].text
            if 'diagnosis' in diagn:
                break
        sub = re.search('diagnosis is (.*?)\.', diagn)
        if sub is not None:
            subStr = sub.group(1)
            ar = re.split(',', subStr)
            allDiagn.add_paragraph(files + '\n' + '|'.join(ar))
allDiagn.save("All The Diagnosis.docx")
