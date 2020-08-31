import tkinter as tk
import tkinter.messagebox
import tkinter.filedialog as fl
import pandas as pd 
from datetime import datetime
from datetime import date

file_path = "C:/Users/PC_1548/Desktop/35.Hafta_Haftalik_Rapor_v2_2020_07_27 - Kopya.xlsx"

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

print("xx" + file_path)

states = pd.read_excel(file_path)
data = states.values.tolist()
print(data[3][5:10])


estimate =  data[3][9]  #data[i][9]
print(estimate )


listOfStrings = estimate.split(" ")
totalWorkDays = 0
for st in listOfStrings:
    if st[-1] == "w":
        totalWorkDays += int(st[:-1]) * 7
    elif st[-1] == "d":
        totalWorkDays += int(st[:-1])
    elif st[-1] == "h":
        totalWorkDays += 1

print(totalWorkDays)

kalanİs = totalWorkDays * (1 - float(data[3][5]))

print(kalanİs)

#print(type(data[3][8]))
endDate = data[3][8] #datetime.strptime(data[3][8], '%Y-%m-%d')
print(endDate)
today = datetime.today() #.strptime("%Y-%m-%d")
print(today)

kalanGunSayisi = (endDate - today).days + 2

#if 

if kalanGunSayisi >= totalWorkDays:
    print("Zamanında")
else:
    print("Riskli")





