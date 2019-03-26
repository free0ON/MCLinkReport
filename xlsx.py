from openpyxl import load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.comments import Comment
from openpyxl.styles import Font
from openpyxl.worksheet.table import Table, TableStyleInfo

def Set_Cell(_wb, name, value):
    range = wb.defined_names[name]
    # if this contains a range of cells then the destinations attribute is not None
    dests = range.destinations  # returns a generator of (worksheet title, cell range) tuples
    for title, coord in dests:
        _ws = _wb[title]
    _ws[coord].value = value

wb = load_workbook(filename=r'C:\Doc\prog\MCLinkReport\templates\Протокол поверки_клин.xlsx', read_only=False)
ws = wb['Лист1']
Set_Cell(wb, 'DocNum', 'B00123/12380823')
Set_Cell(wb, 'EndDate', '12.12.2012')
Set_Cell(wb, 'ReestrNum', '12343-18')
Set_Cell(wb, 'CustomerName', 'ФБУ Клинский ЦСМ')
Set_Cell(wb, 'Range', '1г - 500г')
Set_Cell(wb, 'SerialNumber', '121234234')
Set_Cell(wb, 'Class', 'F1')
Set_Cell(wb, 'TempAvr', '21,5')
Set_Cell(wb, 'HymAvr', '40')
Set_Cell(wb, 'PressAvr', '991')
Set_Cell(wb, 'DensityAvr', '1,15342')
Set_Cell(wb, 'Method', 'МП 17002')
Set_Cell(wb, 'EtalonInfo', '2.1.ZТТ.1956.2017')
Set_Cell(wb, 'Cell1', '123,1234')



wb.save(r'C:\Doc\prog\MCLinkReport\templates\1.xlsx')