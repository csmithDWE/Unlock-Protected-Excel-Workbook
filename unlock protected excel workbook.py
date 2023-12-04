import tkinter as tk
from tkinter.filedialog import askdirectory
from tkinter import filedialog
import os
from zipfile import ZipFile
import xml.etree.ElementTree as ET
import re
import shutil
# path = askdirectory(title = 'Select Folder')
# print(path)

#selects the file and turns it into a .zip instead
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
print(file_path)
file_path_zip = file_path.replace(".xlsx", ".zip")
file_dir = file_path.replace(".xlsx", "")
os.replace(file_path, file_path_zip)

#extracts all for easier file work 
with ZipFile(file_path_zip, 'r') as zip:
    zip.printdir()
    zip.extractall(path = file_dir)
    print('Extraction complete')
    zip.close()

#changes .xml to .txt for easier find/replace
for filename in os.listdir(file_dir + "/xl/worksheets"):
  if filename.endswith('.xml'):
     print(filename)
     file_path_txt = filename.replace(".xml", ".txt")
     os.rename(file_dir + "/xl/worksheets/"+filename, file_dir + "/xl/worksheets/"+file_path_txt)

#does a RegEx find/replace based on sheetPro (replace w nothing)
for filename in os.listdir(file_dir + "/xl/worksheets"):
  if filename.endswith('.txt'):
    print(filename)
    with open (file_dir + "/xl/worksheets/"+filename, 'r' ) as f:
        content = f.read()
        content_new = re.sub('<sheetPro[^>]*>', '', content)
    with open (file_dir + "/xl/worksheets/"+filename, 'w') as f:
       f.write(content_new)
    
#change back from .txt to .xml on all files 
for filename in os.listdir(file_dir + "/xl/worksheets"):
  if filename.endswith('.txt'):
     print(filename)
     file_path_txt = filename.replace(".txt", ".xml")
     os.rename(file_dir + "/xl/worksheets/"+filename, file_dir + "/xl/worksheets/"+file_path_txt)
    
#zip the modded worksheets folder and rename
shutil.make_archive(file_dir+" UNLOCKED", 'zip', file_dir)
unlocked_file = file_dir+" UNLOCKED.zip"
os.rename(unlocked_file, file_dir+ " UNLOCKED.xlsx")

#Delete the files that were used for creation of the new .xlsx (Cleanup)
shutil.rmtree(file_dir)
os.remove(file_dir+".zip")




