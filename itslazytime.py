from docx import Document
import re
import os
import tkinter as tk
from tkinter import filedialog, Text

root = tk.Tk()

canvas = tk.Canvas(root, height = 300, width = 300, bg='#5BCF8C')
canvas.pack()


def runIt():
    allDiagn = Document()
    for files in os.listdir():
        if files.endswith('docx'):
            if files != 'All The Diagnosis.docx':
                doc = Document(files)
                diagn = ''
                for num in range (3):
                    diagn = doc.paragraphs[num].text
                    if 'diagnosis' in diagn:
                        break
                sub = re.search('diagnosis of (.*?)\.', diagn)
                if sub is not None:
                    subStr = sub.group(1)
                    ar = re.split(',', subStr)
                    allDiagn.add_paragraph(files + '\n' + '|'.join(ar))
    allDiagn.save("All The Diagnosis.docx")


runAll = tk.Button(root, text='Run For All Documents', command=runIt)
runAll.pack()

root.mainloop()