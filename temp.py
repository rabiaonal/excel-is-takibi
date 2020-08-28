import pandas as pd

states = pd.read_excel('C:/Users/PC_1548/Downloads/35.Hafta_Haftalik_Rapor_v2_2020_07_27.xlsx')

myList = states.values.tolist()

for row in myList:
    print(row[5])
    
print("\n")
for row in myList:
    print(type(row[8])) # tarihler string


