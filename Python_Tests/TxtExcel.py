from openpyxl import load_workbook
import decimal

def TxtExl():

  #Abre un archivo txt, r = Read
  file = open('src/hsbc.txt', 'r')
  f = file.readlines()

  # Abre un Excel YA creado
  file_path = 'src/Prueba.xlsx'
  wb = load_workbook(file_path)
  ws = wb['Hoja1']  # or wb.active

  #Guarda contenido del Txt en un List

  newList = []
  for line in f:
    if line[-1] == '\n':
      newList.append(line)
    else:
      newList.append(line)  

  #Iteración en la newList e Inserción a Excel

  for index, item in enumerate(newList):

    # Variables
    column = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    #Identifica el tipo de Tag
    #"Switch"
    #Caso Tag 20
    if(item.startswith(":20:")):
      row = 1
      col = 1
      tag = item
      index = [4, 14]
      parts = [tag[i:j] for i,j in zip(index, index[1:]+[None])]
      # Inserción de cada Split del Item por Celda predefinida
      for id, items in enumerate(parts):
          # Si la Fila y Columna tienen Valor, se saltará una Fila para insertar Datos
          if ws.cell(row).value:
            row+=1
          ws[column[id] + str(row)] = items
          #ws['A' + str(2 + row)] = parts[0]
          #ws['B' + str(2 + row)] = parts[1]

    #Caso Tag 25
    elif(item.startswith(":25:")):
      row += 1
      tag = item.split(":25:", 1)
      item = tag[1]
      ws['B' + str(row)] = item
    #Caso Tag 28C
    elif(item.startswith(":28C:")):
      tag = item.split(":28C:", 1)
      aux = tag[1].split("/")
      item = aux[0]
    #Caso Tag 60F
    elif(item.startswith(":60F:")):
      tag = item
      index = [5, 6, 12, 15, 30]
      parts = [tag[i:j] for i,j in zip(index, index[1:]+[None])]
      item = parts[0]
    #Caso Tag 61 // Arreglar el salto de Línea
    elif(item.startswith(":61:")):
      tag = item
      index = [4, 10, 14, 15, 30, 34, 70]
      parts = [tag[i:j] for i,j in zip(index, index[1:]+[None])]
      item = parts[0]
    #Caso Tag 86
    elif(item.startswith(":86:")):
      tag = item.split(":86:", 1)
      item = tag[1]
    #Caso Tag 62a
    elif(item.startswith(":62M:")):
      tag = item
      index = [5, 6, 12, 15, 30]
      parts = [tag[i:j] for i,j in zip(index, index[1:]+[None])]
      item = parts[0]
    else:
      item = ""

  print("Excel Created")
  
  wb.save(file_path)
