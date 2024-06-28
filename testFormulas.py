#Se importa la libreria openpyxl
import openpyxl

############################################################################################
#############################   LO MAS BASICO FORMULAS  #####################################

#Se carga el libro de trabajo
book = openpyxl.load_workbook("ExcelTestEscritura.xlsx")

#Se selecciona la hoja de cálculo activa
sheet = book.active

#Para realizar una formula, se hace igual que en excel
sheet["E1"] = "suma total"
sheet["E2"] = "=SUM(B2:B14)"

#Guardamos el libro de trabajo
book.save("ExcelTestEscritura.xlsx")