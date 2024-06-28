#Se importa la libreria openpyxl
import openpyxl
from openpyxl.chart import ScatterChart, Reference, Series


############################################################################################
#############################   LO MAS BASICO GRAFICAS #####################################

#Se carga el libro de trabajo
book = openpyxl.load_workbook("ExcelTestEscritura.xlsx")
sheet = book.active
sheet4 = book.create_sheet("hoja_4")

for i in range(1,15):
    sheet4[f"A{i}"] = i**2

for i in range(1,15):
    sheet4[f"B{i}"] = i

#Como crear una grafica de dispersion
c1 = ScatterChart()
c1.title = "Grafica de dispersion"
c1.style = 13
c1.y_axis.tittle = "Eje Y"
c1.x_axis.tittle = "Eje X"

xvalues = Reference(sheet4, min_col=1, max_col = 1, min_row=1, max_row=14)
yvalues = Reference(sheet4, min_col=2, max_col = 2, min_row=1, max_row=14)

ser = Series(yvalues, xvalues, title = "recta")

c1.series.append(ser)

sheet4.add_chart(c1, "D3")

book.save("ExcelTestEscritura.xlsx")