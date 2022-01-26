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

"""This program is to pull certain  photos from a selected pool according to excel sheet.
Photos will be copied to a directory of your choice, and will be seperated to different 
folders with sheet names."""

print(" "* 2700)
print("Welcome to PhotoCarry") 
print("This program is written for the purpose of copying photos")
print("from a pool to a directory according to an excelfile") 
print("You will be asked to choose the pool directory  first and then")
print("the excel file. Lastly the location of the new directory.")
print(" "* 900)
#input('Press any key to continue')

root = tk.Tk()
root.withdraw()


#Pop-up that asks from where the  photos are taken 
pool_path = filedialog.askdirectory(title="Choose the location of the photo pool")

#pop-up that asks where the excel file is.
file_path = filedialog.askopenfilename(title="Choose the excel file")
print(file_path)

#Pop-up that asks Where it will be stored
directory_path = filedialog.askdirectory(title="Choose the directory where the photos will be copied to")

file = panda.ExcelFile(file_path)
sheets = file.sheet_names

wb = load_workbook(file_path)


#iterating through the sheets
for sheet in sheets:
    page=file.parse(sheet)
    lenght, widht = page.shape
    print("The lenght of this sheet is :"+ str(lenght))
    ws = wb[sheet]
    path = os.path.join(directory_path, sheet)
    
    #if a folder with sheet's name exist it will fo nothing, else it will create a new one.
    if sheet in os.listdir(directory_path):
        continue
    else:
        os.mkdir(path)
        print("Directory '%s' created" %sheet)

    store_folder = directory_path+"/"+sheet
    for i in range(0, lenght):
        photo_check = page['Inv. No.'][i]
        
        

        #checking if there is a match
        #The algorithm for searching the filename is created according the need of the developer.
         #example of a filename for this algorithm is "ARK_123_4567.jpg" and we want to match with "4567".
        for filename in os.listdir(pool_path):

            #If your photo's name and the cell have the same name simply delete the codes until two lines above the Eureka part.
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
                                    
                                    #This means that the photo's name is the same with the cell.
                                    if shm_number == photo_check:
                                        print("Eureka!!")
                                        shutil.copy(pool_path+"/"+filename, store_folder)
