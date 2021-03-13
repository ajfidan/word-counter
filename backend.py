import pandas as pd
import re
from pandas import read_excel
from collections import Counter, Counter
from tkinter import filedialog

from nltk.corpus import stopwords

#Set exclusion words
stop_words = set(stopwords.words('english'))
stop_words.update('and','i','a','so','arnt', "aren't",'this','when','it','many','so','cant', "can't",'yes','no','these')

#Select Excel File and Column to Read
#Arguments are the name of the sheet and name of column of interest
def SelectColumn(sheetName, columnName):
    filename = filedialog.askopenfilename()
    df = pd.read_excel(filename, sheetName)
    column = df[columnName]
    print('Reading Excel File...')
    return column

#Combine all the cells in the column into one string and then split on punctuaction and white space
#After that, count the frequency of each word excluding the list of stopwords
def SplitColumn(column):
    store = ""
    for i in column.index:
        store += " " 
        store += str(column[i])

    words = re.split(r'[^A-Za-z0-9]+',store.lower())

    filteredWords = [word for word in words if word not in stop_words]
    numberOfWords = Counter(filteredWords)
    return numberOfWords

#Input the dictionary of words and counts and write that to a CSV file
def WriteToCSV(numberOfWords):
    frame=pd.DataFrame.from_dict(numberOfWords, orient='index').reset_index()
    frame.to_csv('../output.csv')
    print('Words saved in output.csv')
