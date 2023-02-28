# importing required modules
import PyPDF2
import pandas as pd
import re
import os

import tkinter as tk
from tkinter import filedialog, messagebox

import sv_ttk

from decimal import Decimal, ROUND_HALF_UP


class GUI(object):
    def __init__(self):
        self.path = ""
        self.root = tk.Tk()

        self.root.title("PDF gradecard to Excell GUI")

        self.label = tk.Label(self.root, text="PDF gradecard to Excell GUI", font=("Times New Roman", 16))
        self.label.grid(row=1, column=2)

        self.textbox = tk.Text(self.root, height=5, font=('Times New Roman', 16))

        self.label = tk.Label(self.root,
                              text=r"Input your PDF file path, like 'C:\Users\(User)\Downloads\karakterutskrift.pdf'",
                              font=("Times New Roman", 12))
        self.label.grid(row=2, column=2)

        self.ent1 = tk.Entry(self.root, font=40)
        self.ent1.grid(row=3, column=2)

        self.b1 = tk.Button(self.root, text="DEM", font=("Times New Roman", 16), command=lambda: self.browsefunc())
        self.b1.grid(row=3, column=3)

        # lambda here because its calling the function instead of passing the function as a callable
        self.b2 = tk.Button(self.root, text="execute", font=("Times New Roman", 16),
                            command=lambda: self.gradeCardtoXLSX(self.path))
        self.b2.grid(row=4, column=2)

        #Set theme
        sv_ttk.set_theme("dark")

        self.root.mainloop()

    def browsefunc(self):
        filename = filedialog.askopenfilename(filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))
        self.ent1.insert(tk.END, filename)  # add this
        print(filename)
        self.path = filename

    def gradeCardtoXLSX(self, pdfLink):

        assert os.path.exists(pdfLink), "File not found at, " + str(pdfLink)
        assert [pdfLink[len(pdfLink) - 4::1] == '.pdf'][0], "The file is not an .pdf at, " + str(pdfLink)
        # creating a pdf file object
        pdfFileObj = open(pdfLink, 'rb')

        # creating a pdf reader object
        reader = PyPDF2.PdfReader(pdfFileObj)

        # creating a page object
        pageObj = reader.pages[0]

        # extracting text from page
        pageText = pageObj.extract_text()

        # get grades
        gradePossible = ("10A", "10B", "10C", "10D", "10E"
                         , "5A", "5B", "5C", "5D", "5E"
                         , "7.5A", "7.5B", "7.5C", "7.5D", "7.5E")
        grade2Num = {
            "A": 6,
            "B": 5,
            "C": 4,
            "D": 3,
            "E": 2
        }

        # Stuff
        gradeList = []
        courseCodes = []

        pattern = "^[A-Z0-9]*$"
        allLines = pageText.split('\n')

        for i, line in enumerate(allLines):
            if line.endswith(gradePossible):
                gradeList.append(grade2Num[line[-1]])
                splitLine = line.split()
                if bool(re.match(pattern, splitLine[0])):
                    courseCodes.append(splitLine[0])
                elif bool(re.match(pattern, allLines[i - 1].split()[0])):
                    courseCodes.append(allLines[i - 1].split()[0])

        meanGrade = [sum(gradeList) / len(gradeList)]
        roundMeanGrade = Decimal(meanGrade[0]).quantize(0, ROUND_HALF_UP)

        # to dataframe
        d = {'Course code': courseCodes + ["Average", "RoundGrade"],
             'grades': gradeList + meanGrade + [list(grade2Num.keys())[list(grade2Num.values()).index(roundMeanGrade)]]}
        df = pd.DataFrame(data=d)

        # closing the pdf file object
        pdfFileObj.close()

        report_path = 'Report Cards'
        if not os.path.exists(report_path):
            os.makedirs(report_path)

        df.to_excel("Report Cards/grades.xlsx")

        #Ask to terminate after succesfull task
        if messagebox.askyesno(title="Success", message="Task successful, quit?"):
            self.root.destroy()


if __name__ == "__main__":
    GUI()
