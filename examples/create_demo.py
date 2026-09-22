from pathlib import Path
from openpyxl import Workbook
root = Path('demo')
root.mkdir(exist_ok=False)
for name, rows in [('january.xlsx', [('Notebook', 10), ('Pen', 20)]),
                   ('february.xlsx', [('Pen', 20), ('Folder', 5)])]:
    book = Workbook()
    book.active.append(['Product', 'Quantity'])
    for row in rows:
        book.active.append(row)
    book.save(root / name)
    book.close()
print('Created two sample workbooks; 3 unique data rows')
