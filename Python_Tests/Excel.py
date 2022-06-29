import xlsxwriter
import decimal

def TxtExl():
  #Abre un archivo txt, r = Read
  file = open('src/hsbc.txt', 'r')
  f = file.readlines()

  #Creación del Excel
  workbook = xlsxwriter.Workbook('src/Prueba.xls')
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

  row = 1
  column = 0
  count = 1 
  #Iteración en la newList e Inserción a Excel
  for item in newList:
    worksheet.write(row, column, item)
    column += 1
    espacio = decimal.Decimal(column)/decimal.Decimal(3)
    if(espacio == count):
      row += 1
      column = 0
      count += 1

  workbook.close()
