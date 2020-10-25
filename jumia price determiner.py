import xlrd
#selling -> profit

path = "C:/Safe/Duala/Afro Kiosk/Documents/Afro Kiosk.xlsx"

def openfile(path):
    book = xlrd.open_workbook(path)
    sheet = book.sheet_by_name("Product list")
    return sheet


def seek(sheet):
    print(sheet.cell_value(6,8)) #column,row
    #product name (6,1) cost price (6,5) jumia price (6,8)

def jumiaProfit(commission,device_cost, delivery, contributions):
    profit = (100 - commission) - device_cost - delivery - contributions
    return profit



sheet = openfile(path)
seek(sheet)