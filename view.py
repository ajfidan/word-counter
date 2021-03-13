import tkinter as tk
from backend import SelectColumn, SplitColumn, WriteToCSV
import pandas as pd
import re
from pandas import read_excel
from collections import Counter, Counter
from tkinter import filedialog

from nltk.corpus import stopwords

#Main function
def Process():
    column = SelectColumn(sheetName.get(), columnName.get())
    numberOfWords = SplitColumn(column)
    WriteToCSV(numberOfWords)

#Windows portion
root = tk.Tk()
root.title('Excel Column Word Frequency Counter')
#root.geometry("600x500")

sheet = tk.StringVar()
column = tk.StringVar()

#Sheet View
tk.Label(root, text='Enter the name of the Excel Sheet: ').grid(row=0)
sheetName = tk.Entry(root, textvar=sheet)
sheetName.grid(row=0, column=1)

#Column View
tk.Label(root, text='Enter the name of the Column: ').grid(row=1)
columnName = tk.Entry(root, textvar=column)
columnName.grid(row=1, column=1)

#Button to read Excel file
tk.Label(root, text='Click the button below to select a file: ').grid(row=5)
tk.Button(root, text='Open File', command=Process).grid(row=6)

root.mainloop()
