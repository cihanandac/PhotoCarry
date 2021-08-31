import pandas as panda
from openpyxl import Workbook
from openpyxl.drawing.image import Image
import os
import shutil
import os.path
import openpyxl
from openpyxl import load_workbook
from openpyxl.drawing.image import Image
import tkinter as tk
from tkinter import filedialog

print(" "* 2900)
print("Welcome to PhotoCarry") 
print("This program is written for the purpose of copying photos")
print("from a pool to a directory according to an excelfile") 
print("        ")
print("Feel free to use and share.")
print("Made by Cihan Andac")
print(" "* 1000)
print("You will be asked to choose the pool directory  first and then")
print("the excel file. Lastly the location of the new directory.")
print(" "* 900)
#input('Press any key to continue')

root = tk.Tk()
root.withdraw()


#Where photos are taken 
pool_path = filedialog.askdirectory(title="Choose the location of the photo pool")

#Where the Excel file is
file_path = filedialog.askopenfilename(title="Choose the excel file")
print(file_path)

#Where it will be stored
directory_path = filedialog.askdirectory(title="Choose the directory where the photos will be copied to")

file = panda.ExcelFile(file_path)
sheets = file.sheet_names

wb = load_workbook(file_path)



for sheet in sheets:
    page=file.parse(sheet)
    lenght, widht = page.shape
    print("The lenght of this sheet is :"+ str(lenght))
    ws = wb[sheet]
    path = os.path.join(directory_path, sheet)
    
    if sheet in os.listdir(directory_path):
        continue
    else:
        os.mkdir(path)
        print("Directory '%s' created" %sheet)

    store_folder = directory_path+"/"+sheet
    for i in range(0, lenght):
        photo_check = page['Inv. No.'][i]
        
        

        #checking if there is a match
        for filename in os.listdir(pool_path):

            
            first_line=0
            
            for k in range(0,len(filename)):
                if filename[k] == "_":
                    if first_line==0:
                        first_line=1
                    else:
                        second_line=0
                        for j in range(0,len(filename)):
                            if filename[j] == "_" or filename[j]== ".":
                                if second_line ==0 or second_line == 1:
                                    second_line = second_line + 1

                            
                                elif second_line ==2:
                                    shm_number = "SHM "+ filename[k+1:j]
                                    

                                    if shm_number == photo_check:
                                        print("Eureka!!")
                                        shutil.copy(pool_path+"/"+filename, store_folder)