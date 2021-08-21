from openpyxl import *
import random as rand
from random import *

def work(wbname, savename, sheetname, outcol):
    wb = load_workbook(wbname)
    ws = wb[sheetname]
    rowc = 1

    for row in ws.rows:
        counter = 0
        for cell in row:
            col = str(cell.fill.start_color.index)
            if col == "FF00FF00":
                counter+=1
        write = ws.cell(row = rowc, column = outcol)
        write.value = counter
        rowc += 1


    wb.save(savename)

work("2021c.xlsx", "counted.xlsx", "do", 10)