import os
import xlwings as xw
from xlwings.base_classes import Book, Table, Sheet, Range

categorias = ['PERFURAÇÃO', 'SÍSMICA', 'GG', 'SSMA']
folder = os.getcwd() + '/Documentos/Indenizáveis/Textos/'
palavras = ['Petra', 'Poço', 'Alumínio']

wb:Book = xw.Book("planilha.xlsx")
ws: Sheet = wb.sheets[0]
table: Table = ws.tables[0]
lastrow = ws.range("A3").end('down').row
lastcolumn = ws.range("A3").end('right').column
coluna_categoria=7
for col in range(1, lastcolumn + 1):
    if ws.range(3, col).value == 'Cod categoria':
        coluna_categoria = col

for i in range(1, len(palavras) + 1):
    ws.range(3, lastcolumn + i).value = palavras[i-1]


for i in range(3, lastrow + 1):
    if ws.range(i, coluna_categoria).value in categorias:
        ws.range(i, 12).value = ["Testei 1", "Testei 2", "Testei 3"]

wb.save("output.xlsx")
wb.close()
