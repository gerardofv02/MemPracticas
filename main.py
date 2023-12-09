import pandas as pd
import random

data = pd.read_excel("./Necesario.xlsx")

correo = "Aqui va su correo"
names = ["Alberto", "Juan", "Pedro", "Miguel", "Ana", "Maria", "Isabel"]


for i in range(len(data)):
    data["Nota"][i] = random.randint(0,10)
    data["Edad"][i] = random.randint(18,25)
    x = random.randint(0,6)
    data["Nombre"][i] = names[x]
    data["Correo"][i] = correo
    

data.to_excel("./Necesario.xlsx","Hoja1",index = False)
