import xlsxwriter
import decimal

def TxtExl():

  # Crear un Excel desde 0 con xlsxwriter

  #Abre un archivo txt, r = Read
  file = open('src/hsbc.txt', 'r')
  f = file.readlines()

  #Creación del Excel
  workbook = xlsxwriter.Workbook('src/Prueba.xlsx')
  worksheet = workbook.add_worksheet("Hoja1")

  #Creación de los Headers
  worksheet.write(0, 0, "Campo1")
  worksheet.write(0, 1, "Campo2")
  worksheet.write(0, 2, "Campo3")

  #Guarda contenido del Txt en un List

  newList = []
  for line in f:
    if line[-1] == '\n':
      newList.append(line)
    else:
      newList.append(line)  

  #Iteración en la newList e Inserción a Excel

  for item in newList:
    row = 1
    
    #Identifica el tipo de Tag
    tag_20 = item.startswith(":20:")
    tag_25 = item.startswith(":25:")
    tag_28C = item.startswith(":28C:")
    tag_60F = item.startswith(":60F:")
    tag_61 = item.startswith(":61:")
    tag_86 = item.startswith(":86:")
    tag_62M = item.startswith(":62M:")

    #"Switch"
    #Caso Tag 20
    if(tag_20):
      # row = 1
      column = 0
      count = 0
      tag = item
      index = [4, 14]
      parts = [tag[i:j] for i,j in zip(index, index[1:]+[None])]
      for item in enumerate(parts):
        worksheet.write(row, column, item[1])
        column += 1
        espacio = decimal.Decimal(column)/decimal.Decimal(1)
        if(espacio == count):
          row += 1
          column += 1
          count += 1
    #Caso Tag 25
    elif(tag_25):
      row += 1
      column = 0
      count = 0
      tag = item.split(":25:", 1)
      item = tag[1]
      worksheet.write(row, column, item)
      column += 1
      espacio = decimal.Decimal(column)/decimal.Decimal(1)
      if(espacio == count):
        column = 0
        count += 1
    #Caso Tag 28C
    elif(tag_28C):
      tag = item.split(":28C:", 1)
      aux = tag[1].split("/")
      item = aux[0]
    #Caso Tag 60F
    elif(tag_60F):
      tag = item
      index = [5, 6, 12, 15, 30]
      parts = [tag[i:j] for i,j in zip(index, index[1:]+[None])]
      item = parts[0]
    #Caso Tag 61 // Arreglar el salto de Línea
    elif(tag_61):
      tag = item
      index = [4, 10, 14, 15, 30, 34, 70]
      parts = [tag[i:j] for i,j in zip(index, index[1:]+[None])]
      item = parts[0]
    #Caso Tag 86
    elif(tag_86):
      tag = item.split(":86:", 1)
      item = tag[1]
    #Caso Tag 62a
    elif(tag_62M):
      tag = item
      index = [5, 6, 12, 15, 30]
      parts = [tag[i:j] for i,j in zip(index, index[1:]+[None])]
      item = parts[0]
    else:
      item = ""

    # worksheet.write(row, column, item)
    # column += 1
    # espacio = decimal.Decimal(column)/decimal.Decimal(1)
    # if(espacio == count):
    #  row += 1
    #  column = 0
    #  count += 1

  print("Excel Created")
  workbook.close()
