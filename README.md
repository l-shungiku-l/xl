# sgpyxl --interface of openpyxl--
## Concept
make typical operaion of Excel file(.xlsx) easier

## Example

```python
import sgpyxl as xl

filepath = "example.xlsx"
wb = xl.SGBook(filepath)
ws = wb.sheet("Sheet1")

list1 = [1, 2, 3]
ws.write_col(list1, 2, 2)

ws.cell(1, 2).value = 0
list2 = ws.get_col(1, 2)

print(xl.recursively_get_file(dirname))

wb.save_and_close()
```
