import tkinter as tk
import tkinter.filedialog as fl
import pandas as pd 
from datetime import datetime
import shutil 
import openpyxl



def helloCallBack():
   global file_path
   file_path = fl.askopenfilename(initialdir = "/",title = "Dosya Seç",filetypes=[("Excel files", "*.xlsx")])
   print(file_path)


window = tk.Tk()
greeting = tk.Label(text="Hello, Tkinter")
greeting.pack()
button = tk.Button(
    text="Dosya Seç",
    width=25,
    height=5,
    bg="blue",
    fg="white",
    command = helloCallBack
)

button.pack()
window.mainloop()


states = pd.read_excel(file_path)
data = states.values.tolist()

estimate = []
for i in range(len(data)):
    
    estimate.append(data[i][9])
 
print(estimate)


totalWorkDays = []
for st in estimate:
    if type(st) == str:
        WorkDays = 0
        listOfStrings = str(st).split()
        for s in listOfStrings:
            if s[-1] == "w":
                WorkDays += int(s[:-1]) * 7
            elif s[-1] == "d":
                WorkDays += int(s[:-1])
            elif s[-1] == "h":
                WorkDays += 1
        totalWorkDays.append(WorkDays)
    else:
        totalWorkDays.append(0)
        
print(totalWorkDays)

kalanIs = []
for i in range(len(data)):
    kalanIs.append(totalWorkDays[i] * (1 - float(data[i][5])))
    
print(kalanIs)

endDate = []
for i in range(len(data)):
    endDate.append(datetime.strptime(data[i][8] , '%Y-%m-%d'))
    

today = datetime.today()
print(today)

kalanGunSayisi = []

for i in range(len(data)):
    kalanGunSayisi.append((endDate[i] - today).days + 2)
    
print(kalanGunSayisi)

file_path_copy = file_path[:-5] + "_CopyFile.xlsx"
shutil.copyfile(file_path, file_path_copy)

theFile = openpyxl.load_workbook(file_path_copy)
currentSheet = theFile[theFile.sheetnames[0]]


for i in range(len(kalanGunSayisi)):
    if kalanGunSayisi[i] >= totalWorkDays[i]:
        currentSheet.cell(i+2 , 7).value = "ZAMANINDA"
        currentSheet.cell(i+2 , 7).fill = openpyxl.styles.PatternFill(bgColor="0000FF00" , fill_type = "solid")
        
    else:

       currentSheet.cell(i+2 , 7).value = "RİSKLİ"
       currentSheet.cell(i+2 , 7).fill = openpyxl.styles.PatternFill(bgColor = "00FFFF00" , fill_type = "solid")
        
theFile.save(file_path_copy)
