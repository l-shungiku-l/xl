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
list2 = []
list2 = ws.get_col(2, 2)

wb.save_and_close()
```
