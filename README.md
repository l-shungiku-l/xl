# sgpyxl --interface of openpyxl--

## Concept

make typical operaion of Excel file(.xlsx) easier

## Example

```python
import sgpyxl as xl

wb = xl.SGBook("./example/example.xlsx")
print(wb.get_sheet_list())
ws = wb.sheet("Sheet1")

# result
# $ ['Sheet1', 'Sheet2', 'Sheet3']
```

```python
list_2d = [["a", 2, 3], [4, "b", 6], [7, 8, "c"]]
ws.write_list_2d(list_2d, 1, 1)
for i in range(1, 4):
    print(ws.get_row(i, 1))

# result
# $ ['a', 2, 3]
#	[4, 'b', 6]
#	[7, 8, 'c']
```

```python
list1 = [1, 4, 7]
list2 = [4, 5, 6]
ws.write_col(list1, 1, 1)
ws.write_row(list2, 2, 1)
ws.cell(3, 3).value = 9
for i in range(1, 4):
    print(ws.get_row(i, 1))

wb.save_and_close()

# result
# $ [1, 2, 3]
#	[4, 5, 6]
#	[7, 8, 9]
#	./example/example.xlsx has been closed
```
