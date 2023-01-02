#! python3
# Convert Barcode List Docx files into txt list

import os
import docx
import pandas as pd
import re
from datetime import datetime
# get current date and time to include in filename
date = datetime.now().strftime("%m%d%y")

#Todo: write regex to remove anything other than barcodes from the input files

#Specify path and convert it to raw data
pathFolder = input('Current path to use: ')
rawFolder = r"{}".format(pathFolder)

#"""Sets the folder variable as the folder specified, and the files as what's in the folder (and then displays the files in the folder)
folder = rawFolder
files = os.listdir(folder)
files

#Sets up empty list, loops through the .docx files in the previously-specified folder, and appends all "paragraphs" to the empty list
fullText = []
for file in files:
    if file.endswith('.docx'):
        doc = docx.Document(f'{folder}/{file}')
        for para in doc.paragraphs:
            fullText.append(para.text) 
print(fullText)

#Takes the list and converts it to a dataframe, strips out all white spaces and rows that are just ""
df = pd.DataFrame (fullText, columns = ['barcode'])
df['barcode'] = df['barcode'].str.replace(" ","")
df.dropna(inplace=True)
df.drop(df.index[df['barcode'] == ""], inplace=True)
print(df)

#Prints the Dataframe to the original folder as a text file without the index, separated by new lines, with a date stamp
df.to_csv(rawFolder + '\\' + 'barcodelist' + date + '.txt', index=None, sep='\n', mode='w')
