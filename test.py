import sgpyxl as xl

wb = xl.SGBook("./test/test.xlsx")
ws = wb.sheet("Sheet1")

print(ws.get_all_cells())
