#Se importa la libreria openpyxl
import openpyxl

############################################################################################
#############################   LO MAS BASICO LECTURA  #####################################

#Se carga el libro de trabajo
book = openpyxl.load_workbook("ExcelTestEscritura.xlsx")

#Se selecciona la hoja de cálculo activa
sheet = book.active

#Se seleccionan las celdas que se desean leer
a1 = sheet["A1"]
a2 = sheet["A2"]

#Para tomar el valor de las celdas se usa el atributo value
print(a1.value)
print(a2.value)
print(type(a1.value))

sheet2 = book["hoja_2"]
print(sheet2["A1"].value)