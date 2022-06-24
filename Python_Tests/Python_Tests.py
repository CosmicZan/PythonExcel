from tkinter import N


#Prueba para Read de un Txt y pasarlo a un Xls

#Abre un archivo txt, r = Read
file = open('src/Prueba.txt', 'r')
f = file.readlines()

#Guarda contenido del Txt en un List
newList = []
for line in f:
  if line[-1] == '\n':
    newList.append(line)
  else:
    newList.append(line)

#Imprime List
print(newList)